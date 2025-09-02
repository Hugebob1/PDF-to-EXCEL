import pypdf
import sys, os, shutil, datetime
from pathlib import Path

from PySide6.QtCore import Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox, QLineEdit
)
def get_inbox_dir() -> Path:

    base = Path.home() / "Documents" / "TataPDF" / "Inbox"
    base.mkdir(parents=True, exist_ok=True)
    return base

def is_pdf(path: Path) -> bool:
    return path.suffix.lower() == ".pdf"

def unique_dest_path(dest_dir: Path, src_name: str) -> Path:
    return dest_dir / src_name

def copy_pdf(src: Path, dest_dir: Path) -> Path:
    dest = unique_dest_path(dest_dir, src.name)
    shutil.copy2(src, dest)
    return dest

# ----------------- Strefa Drag&Drop -----------------
class DropZone(QLabel):
    filesDropped = Signal(list)

    def __init__(self, text="Upuść PDF tutaj"):
        super().__init__(text)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #888;
                border-radius: 10px;
                padding: 30px;
                font-size: 16px;
                color: #444;
                background: #fafafa;
            }
            QLabel:hover { border-color: #555; }
        """)

    def dragEnterEvent(self, event):
        md = event.mimeData()
        if md.hasUrls() and any(url.isLocalFile() and is_pdf(Path(url.toLocalFile())) for url in md.urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        md = event.mimeData()
        paths = []
        if md.hasUrls():
            for url in md.urls():
                if url.isLocalFile():
                    p = Path(url.toLocalFile())
                    if p.exists() and is_pdf(p):
                        paths.append(str(p))
        if paths:
            self.filesDropped.emit(paths)
        event.acceptProposedAction()

def read_from_pdf(path):
    reader = pypdf.PdfReader(path)


    # print(len(reader.pages))
    if reader.pages[0].extract_text():
        return len(reader.pages[0].extract_text())
    return "Nie udalo sie odczytac pdfa."

# ------------------ Główne okno ---------------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Importer PDF → Excel (schowek plików)")
        self.resize(700, 450)

        self.drop = DropZone()
        self.btn_pick = QPushButton("Wybierz PDF…")
        self.btn_open_folder = QPushButton("Otwórz folder docelowy")
        self.status = QLineEdit()
        self.status.setReadOnly(True)
        self.status.setPlaceholderText("Tu pojawi się status importu…")

        # Layout
        top_row = QHBoxLayout()
        top_row.addWidget(self.btn_pick)
        top_row.addWidget(self.btn_open_folder)

        layout = QVBoxLayout(self)
        layout.addLayout(top_row)
        layout.addWidget(self.drop)
        layout.addWidget(self.status)

        self.drop.filesDropped.connect(self.process_files)
        self.btn_pick.clicked.connect(self.pick_files)
        self.btn_open_folder.clicked.connect(self.open_dest_folder)

        self.dest_dir = get_inbox_dir()

    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Wybierz pliki PDF", str(Path.home()), "PDF Files (*.pdf)"
        )
        if paths:
            self.process_files(paths)

    def process_files(self, paths: list[str]):
        count_ok, count_fail = 0, 0
        saved = []

        for s in paths:
            src = Path(s)
            try:
                if not src.exists():
                    raise FileNotFoundError("Plik nie istnieje")
                if not is_pdf(src):
                    raise ValueError("To nie jest PDF")

                dest = copy_pdf(src, self.dest_dir)
                saved.append(dest.name)
                count_ok += 1
            except Exception as e:
                count_fail += 1

        if count_ok and not count_fail:
            self.status.setText(f"Plik ma {read_from_pdf(dest)} znakow.")
        elif count_ok and count_fail:
            self.status.setText(f"OK: {count_ok}, błędów: {count_fail} — folder: {self.dest_dir}")
        else:
            self.status.setText("Nie udało się zapisać żadnego pliku.")

    def open_dest_folder(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.dest_dir)  # type: ignore[attr-defined]
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.dest_dir)))
        except Exception:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.dest_dir)))

# ------------------- Uruchomienie -------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
