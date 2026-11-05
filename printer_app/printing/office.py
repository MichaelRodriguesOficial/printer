"""Impressão de documentos do Office (Word/Excel via COM) com alternativa no LibreOffice.

Cuidados que a versão 1.0 não tinha:

* a impressora escolhida é aplicada de fato (Word: ``FilePrintSetup`` sem alterar a
  impressora padrão do Windows; Excel: ``ActivePrinter`` no formato "Nome em Ne01:");
* ``PrintOut(Background=False)`` — o ``Quit`` não cancela mais a impressão em andamento;
* nada de diálogos invisíveis travando o processo (``DisplayAlerts`` desligado, abertura
  somente leitura, senha falsa para documentos protegidos, fechar sem salvar);
* macros sempre desabilitadas (``AutomationSecurity = msoAutomationSecurityForceDisable``)
  — via automação o padrão do Office é *habilitar* macros sem perguntar;
* uma única instância de Word/Excel por lote, sempre encerrada em ``finally`` e isolada
  dos documentos que o usuário estiver com abertos (``DispatchEx``).
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

log = logging.getLogger(__name__)

MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE = 0
WD_STATISTIC_PAGES = 2
XL_TYPE_PDF = 0
WD_EXPORT_FORMAT_PDF = 17
# Senha "falsa": em arquivo protegido gera erro em vez de um prompt invisível; em
# arquivos sem senha é ignorada pelo Office.
_DUMMY_PASSWORD = "__printer_cav_sem_senha__"
# Conector localizado de Application.ActivePrinter do Excel ("HP em Ne01:").
_EXCEL_CONNECTORS = ("em", "on", "en", "sur", "auf", "su", "op")


class OfficeError(Exception):
    """Falha ao imprimir documento do Office, com mensagem pronta para o usuário."""


def _com_message(exc: Exception) -> str:
    """Extrai a descrição legível de um ``pywintypes.com_error``."""
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return str(exc)


def _friendly(exc: Exception) -> str:
    text = _com_message(exc)
    lower = text.lower()
    if "senha" in lower or "password" in lower:
        return "Documento protegido por senha"
    return text


def excel_printer_candidates(printer: str) -> list[str]:
    """Nomes aceitos por ``Excel.ActivePrinter`` para a impressora, do mais provável ao menos."""
    port = None
    try:
        import winreg

        key = r"Software\Microsoft\Windows NT\CurrentVersion\Devices"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as handle:
            value, _ = winreg.QueryValueEx(handle, printer)
        port = str(value).split(",")[-1].strip() or None
    except (OSError, ImportError):
        pass
    candidates = [f"{printer} {conn} {port}" for conn in _EXCEL_CONNECTORS] if port else []
    return [*candidates, printer]


class OfficeSession:
    """Mantém Word/Excel abertos (ocultos) durante um lote e os encerra no final.

    Deve ser usada inteiramente dentro da mesma thread (exigência do COM).
    """

    def __init__(self, com_available: bool | None = None) -> None:
        self.com_available = os.name == "nt" if com_available is None else com_available
        self._word = None
        self._excel = None
        self._com_ready = False
        self._libreoffice: str | bool | None = False  # False = ainda não procurado
        self._lo_profile: Path | None = None

    # --- ciclo de vida -------------------------------------------------------

    def __enter__(self) -> OfficeSession:
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.close()

    def close(self) -> None:
        for attr in ("_word", "_excel"):
            app = getattr(self, attr)
            if app is not None:
                try:
                    app.Quit()
                except Exception as exc:  # noqa: BLE001
                    log.warning("Falha ao encerrar %s: %s", attr, exc)
                setattr(self, attr, None)
        if self._com_ready:
            import pythoncom

            pythoncom.CoUninitialize()
            self._com_ready = False
        if self._lo_profile is not None:
            shutil.rmtree(self._lo_profile, ignore_errors=True)
            self._lo_profile = None

    def _dispatch(self, prog_id: str):
        if not self.com_available:
            return None
        import pythoncom
        import win32com.client

        if not self._com_ready:
            pythoncom.CoInitialize()
            self._com_ready = True
        try:
            app = win32com.client.DispatchEx(prog_id)
        except Exception as exc:  # noqa: BLE001 - Office não instalado
            log.info("%s indisponível: %s", prog_id, _com_message(exc))
            return None
        app.Visible = False
        app.DisplayAlerts = WD_ALERTS_NONE if prog_id.startswith("Word") else False
        try:
            app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception as exc:  # noqa: BLE001
            log.warning("Não foi possível desabilitar macros em %s: %s", prog_id, exc)
        return app

    def word(self):
        if self._word is None:
            self._word = self._dispatch("Word.Application")
        return self._word

    def excel(self):
        if self._excel is None:
            self._excel = self._dispatch("Excel.Application")
        return self._excel

    # --- Word ------------------------------------------------------------------

    def print_word(self, path: Path, printer: str, copies: int) -> int | None:
        word = self.word()
        if word is None:
            return None
        try:
            doc = word.Documents.Open(
                FileName=str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=_DUMMY_PASSWORD,
                Visible=False,
                NoEncodingDialog=True,
            )
        except Exception as exc:
            raise OfficeError(_friendly(exc)) from exc
        try:
            self._set_word_printer(word, printer)
            pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            doc.PrintOut(Background=False, Copies=copies, Collate=True)
            return int(pages) * copies
        except OfficeError:
            raise
        except Exception as exc:
            raise OfficeError(_friendly(exc)) from exc
        finally:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE)
            except Exception as exc:  # noqa: BLE001
                log.warning("Falha ao fechar documento %s: %s", path, exc)

    @staticmethod
    def _set_word_printer(word, printer: str) -> None:
        try:
            # Altera só a impressora do Word, sem mexer na padrão do Windows.
            word.WordBasic.FilePrintSetup(Printer=printer, DoNotSetAsSysDefault=1)
            return
        except Exception as exc:  # noqa: BLE001
            log.info("FilePrintSetup falhou (%s); usando ActivePrinter", _com_message(exc))
        import win32print

        previous = win32print.GetDefaultPrinter()
        try:
            word.ActivePrinter = printer  # efeito colateral: muda a padrão do Windows...
        except Exception as exc:
            raise OfficeError(f"O Word não aceitou a impressora: {_com_message(exc)}") from exc
        finally:
            if previous and previous != printer:
                win32print.SetDefaultPrinter(previous)  # ...que restauramos aqui.

    # --- Excel -----------------------------------------------------------------

    def print_excel(self, path: Path, printer: str, copies: int) -> int | None:
        excel = self.excel()
        if excel is None:
            return None
        try:
            wb = excel.Workbooks.Open(
                Filename=str(path),
                UpdateLinks=0,
                ReadOnly=True,
                Password=_DUMMY_PASSWORD,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
        except Exception as exc:
            raise OfficeError(_friendly(exc)) from exc
        try:
            if self._set_excel_printer(excel, printer):
                wb.PrintOut(Copies=copies, Collate=True)
                return None
            # Nenhum formato de nome foi aceito: exporta para PDF e imprime via GDI.
            log.info("Excel recusou a impressora %r; imprimindo via PDF", printer)
            pdf = Path(tempfile.mkdtemp(prefix="printer_xl_")) / (path.stem + ".pdf")
            wb.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=str(pdf))
            raise _PrintViaPdf(pdf)
        except (_PrintViaPdf, OfficeError):
            raise
        except Exception as exc:
            raise OfficeError(_friendly(exc)) from exc
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception as exc:  # noqa: BLE001
                log.warning("Falha ao fechar planilha %s: %s", path, exc)

    @staticmethod
    def _set_excel_printer(excel, printer: str) -> bool:
        for candidate in excel_printer_candidates(printer):
            try:
                excel.ActivePrinter = candidate
                return True
            except Exception:  # noqa: BLE001 - tenta o próximo formato
                continue
        return False

    # --- LibreOffice (alternativa quando não há Microsoft Office) ---------------

    def libreoffice(self) -> str | None:
        if self._libreoffice is False:
            self._libreoffice = find_libreoffice()
        return self._libreoffice or None

    def convert_to_pdf(self, path: Path, outdir: Path) -> Path:
        soffice = self.libreoffice()
        if soffice is None:
            raise OfficeError(
                "Instale o Microsoft Office ou o LibreOffice para imprimir este arquivo"
            )
        if self._lo_profile is None:
            self._lo_profile = Path(tempfile.mkdtemp(prefix="printer_lo_"))
        outdir.mkdir(parents=True, exist_ok=True)
        cmd = [
            soffice,
            f"-env:UserInstallation={self._lo_profile.as_uri()}",
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(path),
        ]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=180, creationflags=flags
            )
        except subprocess.TimeoutExpired as exc:
            raise OfficeError("LibreOffice demorou demais para converter o arquivo") from exc
        pdf = outdir / (path.stem + ".pdf")
        if result.returncode != 0 or not pdf.exists():
            detail = (result.stderr or result.stdout).strip().splitlines()
            raise OfficeError(
                "LibreOffice não conseguiu converter o arquivo"
                + (f" ({detail[-1]})" if detail else "")
            )
        return pdf


class _PrintViaPdf(Exception):  # noqa: N818 - sinal interno, não é erro
    def __init__(self, pdf: Path) -> None:
        super().__init__(str(pdf))
        self.pdf = pdf


def find_libreoffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found
    if os.name == "nt":
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                exe = Path(base) / "LibreOffice" / "program" / "soffice.exe"
                if exe.exists():
                    return str(exe)
    return None
