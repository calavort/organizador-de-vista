"""Integracao Qt nao bloqueante do atualizador do Organizador de Vista.

Vem do Super Captura, sem a parte que preserva a edicao em andamento: aqui
nao ha desenho na tela para recuperar, so uma operacao do Tekla que nao pode
ser interrompida no meio.
"""

from __future__ import annotations

import json
import tempfile
import threading
from pathlib import Path

from PySide6.QtCore import QObject, QTimer, Signal, Slot
from PySide6.QtWidgets import QMessageBox

from .atualizador import (UpdateError, check_release, download_release, prepare_installer,
                          read_version, start_installer, state_path, unpack_package)


class UpdateController(QObject):
    completed = Signal(str, object)
    failed = Signal(str)
    progress = Signal(int)

    def __init__(self, window, root: Path):
        super().__init__(window)
        self.window = window
        self.root = root
        self.info = read_version(root)
        self.release = None
        self.archive = None
        self.busy = False
        self.installing = False
        self.started = False
        self.dialog = None
        self.message = "Aguardando verificacao."
        self.manual = False
        self.helper = None
        self._handshake_attempts = 0
        self.handshake_timer = QTimer(self)
        self.handshake_timer.setInterval(100)
        self.handshake_timer.timeout.connect(self._check_handshake)
        self.completed.connect(self._completed)
        self.failed.connect(self._failed)
        self.progress.connect(self._progress)

    # ------------------------------------------------------------------ apoio
    def _log(self, message):
        """Escreve no registro que aparece na tela do programa."""
        try:
            self.window.controller.log(message)
        except Exception:
            pass

    def _tekla_busy(self):
        """Uma operacao do Tekla em andamento nao pode ser interrompida."""
        try:
            return bool(self.window.bridge.is_busy())
        except Exception:
            return False

    def _status(self, message):
        self.message = message
        payload = {"version": self.info["version"], "message": message,
                   "busy": self.busy, "available": self.release is not None}
        self.window.web.page().runJavaScript(
            "if (typeof updateReleaseState === 'function') updateReleaseState("
            + json.dumps(payload) + ");"
        )

    def _worker(self, operation, callback):
        self.busy = True

        def run():
            try:
                value = callback()
                self.completed.emit(operation, value)
            except Exception as exc:
                try:
                    self.failed.emit(str(exc))
                except RuntimeError:
                    pass  # A janela foi fechada durante a consulta na rede.

        threading.Thread(target=run, name="OrganizadorDeVista-" + operation, daemon=True).start()

    # ------------------------------------------------------------- ciclo de uso
    def ui_ready(self):
        self._status(self.message)
        if self.started:
            return
        self.started = True
        resultado = state_path(self.root) / "resultado.json"
        if resultado.exists():
            try:
                dados = json.loads(resultado.read_text(encoding="utf-8"))
                self._log(dados["message"])
                resultado.replace(resultado.with_name("ultimo-resultado.json"))
            except (OSError, ValueError, KeyError):
                pass
        QTimer.singleShot(1800, lambda: self.check(False))

    def check(self, manual=True):
        if self.busy or self.installing:
            return
        self.manual = manual
        self._worker("check", lambda: check_release(self.info))
        self._status("Verificando atualizacoes...")

    def offer_install(self):
        if not self.release or self.busy or self.dialog:
            return
        if self._tekla_busy() or not self.window.isVisible():
            self._status("Atualizacao disponivel. Conclua a operacao em andamento para instalar.")
            return
        self.dialog = QMessageBox(self.window)
        self.dialog.setWindowTitle("Atualizacao do Organizador de Vista")
        self.dialog.setIcon(QMessageBox.Icon.Information)
        self.dialog.setText(f"A versao {self.release.version} esta disponivel.")
        self.dialog.setInformativeText("Baixar a atualizacao agora? Voce confirma a reinicializacao depois do download.")
        baixar = self.dialog.addButton("Baixar atualizacao", QMessageBox.ButtonRole.AcceptRole)
        self.dialog.addButton("Agora nao", QMessageBox.ButtonRole.RejectRole)
        self.dialog.setDefaultButton(baixar)

        def respondeu():
            aceitou = self.dialog.clickedButton() == baixar
            self.dialog.deleteLater()
            self.dialog = None
            if aceitou:
                self._download()

        self.dialog.finished.connect(respondeu)
        self.dialog.open()

    def _download(self):
        if self.archive and self.archive.exists():
            self.confirm_restart()
            return
        release = self.release

        def work():
            archive = download_release(release, self.root, self.progress.emit)
            try:
                with tempfile.TemporaryDirectory(prefix="validacao-", dir=state_path(self.root)) as temporary:
                    stage = Path(temporary)
                    unpack_package(archive, stage, release.version, release.repository)
                    if (stage / "requirements.txt").read_bytes() != (self.root / "requirements.txt").read_bytes():
                        raise UpdateError("Esta versao altera as bibliotecas. Rode o Instalar dependencias.bat "
                                          "depois de instalar o pacote a mao.")
            except Exception:
                archive.unlink(missing_ok=True)
                raise
            return archive

        self._worker("download", work)
        self._status("Baixando atualizacao...")

    @Slot(str, object)
    def _completed(self, operation, value):
        self.busy = False
        if operation == "check":
            if self.release != value and self.archive:
                self.archive.unlink(missing_ok=True)
                self.archive = None
            self.release = value
            self._status(f"Versao {value.version} disponivel." if value else "Voce esta na versao mais recente.")
            if self.manual:
                self._log(self.message)
            if value:
                self.offer_install()
        else:
            self.archive = value
            self._status("Download conferido. Pronto para instalar.")
            self.confirm_restart()

    @Slot(str)
    def _failed(self, message):
        self.busy = False
        self.installing = False
        self.window.setEnabled(True)
        self._status(message)
        if self.manual or self.archive:
            self._log(message)

    @Slot(int)
    def _progress(self, percent):
        self._status(f"Baixando atualizacao: {percent}%")

    def confirm_restart(self):
        if self._tekla_busy() or not self.window.isVisible():
            self._status("Download pronto. Conclua a operacao em andamento para instalar.")
            return
        self.dialog = QMessageBox(self.window)
        self.dialog.setWindowTitle("Instalar atualizacao")
        self.dialog.setText(f"Instalar a versao {self.release.version} e reiniciar?")
        self.dialog.setInformativeText("O programa fecha e abre sozinho. O Tekla nao e afetado.")
        instalar = self.dialog.addButton("Instalar e reiniciar", QMessageBox.ButtonRole.AcceptRole)
        depois = self.dialog.addButton("Mais tarde", QMessageBox.ButtonRole.RejectRole)
        self.dialog.setDefaultButton(depois)

        def respondeu():
            aceitou = self.dialog.clickedButton() == instalar
            self.dialog.deleteLater()
            self.dialog = None
            if aceitou:
                self._begin_install()

        self.dialog.finished.connect(respondeu)
        self.dialog.open()

    def _begin_install(self):
        if self.installing:
            return
        if self._tekla_busy():
            self._status("Conclua a operacao em andamento antes de instalar.")
            return
        # Mudanca em desenvolvimento se publica, nao se troca por um pacote antigo.
        if (self.root / ".git").exists():
            self._status("Pasta de desenvolvimento. Teste a instalacao na copia extraida do pacote ZIP.")
            self._log(self.message)
            return
        self.installing = True
        self.window.setEnabled(False)
        try:
            plano = prepare_installer(self.archive, self.root, self.release)
            self.helper = start_installer(plano)
            self._handshake_attempts = 0
            self.handshake_timer.start()
        except Exception as exc:
            self._failed(str(exc))

    def _check_handshake(self):
        self._handshake_attempts += 1
        if self.helper.poll() is not None:
            self.handshake_timer.stop()
            self._failed("O instalador nao iniciou. O programa foi mantido aberto.")
        elif (state_path(self.root) / "instalador-pronto").exists():
            self.handshake_timer.stop()
            self.window.close()
        elif self._handshake_attempts >= 100:
            self.handshake_timer.stop()
            self.helper.terminate()
            self._failed("O instalador nao respondeu. Tente novamente.")
