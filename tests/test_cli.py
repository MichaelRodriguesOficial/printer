from __future__ import annotations

import os

import pytest

from printer_app import __version__
from printer_app.__main__ import main


def test_version(capsys):
    with pytest.raises(SystemExit) as info:
        main(["--version"])
    assert info.value.code == 0
    assert __version__ in capsys.readouterr().out


@pytest.mark.skipif(os.name == "nt", reason="no Windows o app abre normalmente")
def test_non_windows_requires_demo(capsys):
    assert main([]) == 2
    assert "--demo" in capsys.readouterr().out


@pytest.mark.skipif(os.name != "nt" and not os.environ.get("DISPLAY"), reason="sem display gráfico")
def test_self_test_passes():
    assert main(["--self-test"]) == 0
