from __future__ import annotations

import shutil
import sys
import types
from pathlib import Path

import pytest

from printer_app.files import Entry, FileKind
from printer_app.printing import PrintContext, PrintError, office, print_entry
from printer_app.printing.office import OfficeError, OfficeSession, _PrintViaPdf


class ComError(Exception):
    """Imita pywintypes.com_error: (hresult, msg, (code, source, descrição, ...), arg)."""

    def __init__(self, description: str) -> None:
        super().__init__(
            -2147352567, "Exceção.", (0, "Microsoft Word", description, None, 0, 0), None
        )


class Recorder:
    def __init__(self, name: str, log: list) -> None:
        object.__setattr__(self, "_name", name)
        object.__setattr__(self, "_log", log)

    def __setattr__(self, key, value):
        self._log.append((f"{self._name}.{key}=", value))
        object.__setattr__(self, key, value)


class FakeDoc:
    def __init__(self, log, fail_print=False):
        self.log, self.fail_print = log, fail_print

    def ComputeStatistics(self, stat):
        return 3

    def PrintOut(self, **kw):
        self.log.append(("PrintOut", kw))
        if self.fail_print:
            raise ComError("A impressora não respondeu")

    def Close(self, **kw):
        self.log.append(("Close", kw))


class FakeWordBasic:
    def __init__(self, log, fail=False):
        self.log, self.fail = log, fail

    def FilePrintSetup(self, **kw):
        self.log.append(("FilePrintSetup", kw))
        if self.fail:
            raise ComError("não suportado")


def make_word(log, *, open_error=None, fail_print=False, fail_setup=False):
    word = Recorder("word", log)
    documents = types.SimpleNamespace()

    def open_(**kw):
        log.append(("Open", kw))
        if open_error:
            raise ComError(open_error)
        return FakeDoc(log, fail_print)

    documents.Open = open_
    object.__setattr__(word, "Documents", documents)
    object.__setattr__(word, "WordBasic", FakeWordBasic(log, fail_setup))
    object.__setattr__(word, "Quit", lambda: log.append(("Quit", "word")))
    return word


@pytest.fixture
def com(monkeypatch):
    """Instala win32com/pythoncom/win32print falsos e devolve o log de chamadas."""
    log: list = []
    apps: dict = {}
    client = types.ModuleType("win32com.client")

    def dispatch_ex(prog_id):
        log.append(("DispatchEx", prog_id))
        if prog_id not in apps:
            raise ComError("Cadeia de caracteres de classe inválida")
        return apps[prog_id]

    client.DispatchEx = dispatch_ex
    pkg = types.ModuleType("win32com")
    pkg.client = client
    pythoncom = types.ModuleType("pythoncom")
    pythoncom.CoInitialize = lambda: log.append(("CoInitialize",))
    pythoncom.CoUninitialize = lambda: log.append(("CoUninitialize",))
    win32print = types.ModuleType("win32print")
    default = {"name": "Padrao"}
    win32print.GetDefaultPrinter = lambda: default["name"]
    win32print.SetDefaultPrinter = lambda n: (
        log.append(("SetDefaultPrinter", n)),
        default.update(name=n),
    )
    for name, mod in (
        ("win32com", pkg),
        ("win32com.client", client),
        ("pythoncom", pythoncom),
        ("win32print", win32print),
    ):
        monkeypatch.setitem(sys.modules, name, mod)
    return types.SimpleNamespace(log=log, apps=apps, default=default)


def test_word_prints_on_selected_printer_safely(com, tmp_path):
    com.apps["Word.Application"] = make_word(com.log)
    with OfficeSession(com_available=True) as session:
        pages = session.print_word(tmp_path / "carta.docx", "BRA3-FINANCEIRO", copies=2)
    assert pages == 6
    log = com.log
    assert ("word.AutomationSecurity=", 3) in log  # macros desabilitadas
    assert ("word.DisplayAlerts=", 0) in log
    assert ("word.Visible=", False) in log
    open_kw = next(e[1] for e in log if e[0] == "Open")
    assert open_kw["ReadOnly"] is True and open_kw["AddToRecentFiles"] is False
    assert open_kw["PasswordDocument"]  # evita prompt invisível de senha
    assert ("FilePrintSetup", {"Printer": "BRA3-FINANCEIRO", "DoNotSetAsSysDefault": 1}) in log
    assert ("PrintOut", {"Background": False, "Copies": 2, "Collate": True}) in log
    assert ("Close", {"SaveChanges": 0}) in log
    assert log[-2:] == [("Quit", "word"), ("CoUninitialize",)]
    assert not any(name == "SetDefaultPrinter" for name, *_ in log)


def test_word_reuses_instance_and_closes_doc_on_error(com, tmp_path):
    com.apps["Word.Application"] = make_word(com.log, fail_print=True)
    with OfficeSession(com_available=True) as session:
        for _ in range(2):
            with pytest.raises(OfficeError, match="não respondeu"):
                session.print_word(tmp_path / "x.docx", "P", 1)
    assert [n for n, *_ in com.log].count("DispatchEx") == 1
    assert [n for n, *_ in com.log].count("Close") == 2
    assert ("Quit", "word") in com.log


def test_word_password_message(com, tmp_path):
    com.apps["Word.Application"] = make_word(com.log, open_error="A senha está incorreta.")
    with OfficeSession(com_available=True) as session, pytest.raises(OfficeError) as info:
        session.print_word(tmp_path / "x.docx", "P", 1)
    assert str(info.value) == "Documento protegido por senha"


def test_word_fallback_restores_default_printer(com, tmp_path):
    com.apps["Word.Application"] = make_word(com.log, fail_setup=True)
    with OfficeSession(com_available=True) as session:
        session.print_word(tmp_path / "x.docx", "Outra", 1)
    assert ("word.ActivePrinter=", "Outra") in com.log
    assert ("SetDefaultPrinter", "Padrao") in com.log
    assert com.default["name"] == "Padrao"


class FakeWorkbook:
    def __init__(self, log):
        self.log = log

    def PrintOut(self, **kw):
        self.log.append(("wb.PrintOut", kw))

    def ExportAsFixedFormat(self, **kw):
        Path(kw["Filename"]).write_bytes(b"%PDF")
        self.log.append(("Export", kw))

    def Close(self, **kw):
        self.log.append(("wb.Close", kw))


class FakeExcel:
    def __init__(self, log, accepted: str | None):
        self.log, self.accepted = log, accepted
        self.Workbooks = types.SimpleNamespace(Open=self._open)

    def _open(self, **kw):
        self.log.append(("wb.Open", kw))
        return FakeWorkbook(self.log)

    def __setattr__(self, key, value):
        if key == "ActivePrinter":
            self.log.append(("try", value))
            if value != self.accepted:
                raise ComError("ActivePrinter inválido")
        object.__setattr__(self, key, value)

    def Quit(self):
        self.log.append(("Quit", "excel"))


def test_excel_uses_localized_active_printer(com, tmp_path, monkeypatch):
    monkeypatch.setattr(
        office, "excel_printer_candidates", lambda p: [f"{p} on Ne02:", f"{p} em Ne02:", p]
    )
    com.apps["Excel.Application"] = FakeExcel(com.log, accepted="HP em Ne02:")
    with OfficeSession(com_available=True) as session:
        assert session.print_excel(tmp_path / "t.xlsx", "HP", 3) is None
    assert ("wb.PrintOut", {"Copies": 3, "Collate": True}) in com.log
    assert ("wb.Close", {"SaveChanges": False}) in com.log
    assert ("Quit", "excel") in com.log


def test_excel_falls_back_to_pdf_when_printer_rejected(com, tmp_path, monkeypatch):
    monkeypatch.setattr(office, "excel_printer_candidates", lambda p: [p])
    com.apps["Excel.Application"] = FakeExcel(com.log, accepted=None)
    with OfficeSession(com_available=True) as session, pytest.raises(_PrintViaPdf) as info:
        session.print_excel(tmp_path / "t.xlsx", "HP", 1)
    assert info.value.pdf.read_bytes() == b"%PDF"
    assert ("wb.Close", {"SaveChanges": False}) in com.log
    shutil.rmtree(info.value.pdf.parent)


def test_excel_candidates_without_registry():
    assert office.excel_printer_candidates("HP")[-1] == "HP"


def test_dispatch_falls_back_to_libreoffice_when_no_office(com, tmp_path, monkeypatch):
    """Sem Word instalado: converte com LibreOffice e imprime o PDF via GDI."""
    printed = []
    monkeypatch.setattr(
        "printer_app.printing._print_pdf",
        lambda path, title, ctx: printed.append((path.suffix, path.exists(), title)) or 1,
    )
    session = OfficeSession(com_available=True)

    def fake_convert(path, outdir):
        pdf = outdir / (path.stem + ".pdf")
        pdf.write_bytes(b"%PDF")
        return pdf

    monkeypatch.setattr(session, "convert_to_pdf", fake_convert)
    doc = tmp_path / "carta.docx"
    doc.write_bytes(b"x")
    ctx = PrintContext("P", office=session)
    assert print_entry(Entry(doc, doc.name, "", FileKind.WORD), ctx) == 1
    assert printed == [(".pdf", True, "carta.docx")]


def test_no_office_and_no_libreoffice(com, tmp_path, monkeypatch):
    monkeypatch.setattr(office, "find_libreoffice", lambda: None)
    doc = tmp_path / "carta.docx"
    doc.write_bytes(b"x")
    ctx = PrintContext("P", office=OfficeSession(com_available=True))
    with pytest.raises(PrintError, match="LibreOffice"):
        print_entry(Entry(doc, doc.name, "", FileKind.WORD), ctx)


@pytest.mark.slow
@pytest.mark.skipif(office.find_libreoffice() is None, reason="LibreOffice não instalado")
@pytest.mark.parametrize(
    ("name", "content"),
    [
        ("comunicado.txt", "Comunicado às consultoras\nLinha 2\n"),
        ("metas.csv", "Consultora;Meta\nAna;1000\nBruna;1500\n"),
    ],
)
def test_real_libreoffice_conversion(tmp_path, name, content):
    from printer_app.printing import pdf

    src = tmp_path / name
    src.write_text(content, encoding="utf-8")
    with OfficeSession(com_available=False) as session:
        out = session.convert_to_pdf(src, tmp_path / "out")
    assert out.suffix == ".pdf"
    assert pdf.page_count(out) >= 1
