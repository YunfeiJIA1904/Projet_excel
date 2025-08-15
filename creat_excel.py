import sys
import json
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QLineEdit, QPushButton, QComboBox, QListWidget, QScrollArea, QMessageBox

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modifier les données")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.dropdown = QComboBox()
        self.dropdown.addItem("AddNew")
        self.dropdown.currentIndexChanged.connect(self.show_frame)
        self.layout.addWidget(self.dropdown)

        self.frames = {}
        self.create_parametre_frame()
        self.load_parameters()
        self.show_frame(0)

    def create_parametre_frame(self):
        frame = QWidget()
        layout = QVBoxLayout(frame)

        self.listbox_headers = QListWidget()
        layout.addWidget(self.listbox_headers)

        self.header_input = QLineEdit()
        self.header_input.setPlaceholderText("Entrez le nom de l'entête")
        layout.addWidget(self.header_input)

        add_header_button = QPushButton("Ajouter Entête")
        add_header_button.clicked.connect(self.add_header)
        layout.addWidget(add_header_button)

        self.tab_name_input = QLineEdit()
        self.tab_name_input.setPlaceholderText("Entrez le nom de l'onglet")
        layout.addWidget(self.tab_name_input)

        save_headers_button = QPushButton("Sauvegarder")
        save_headers_button.clicked.connect(self.save_headers)
        layout.addWidget(save_headers_button)

        self.frames["AddNew"] = frame
        self.layout.addWidget(frame)

    def add_header(self):
        header = self.header_input.text()
        if header and not any(
                header == self.listbox_headers.item(i).text() for i in range(self.listbox_headers.count())):
            self.listbox_headers.addItem(header)
        self.header_input.clear()

    def save_headers(self):
        headers = [self.listbox_headers.item(i).text() for i in range(self.listbox_headers.count())]
        tab_name = self.tab_name_input.text()
        if headers and tab_name:
            self.dropdown.addItem(tab_name)
            new_scroll_area = self.create_scrollable_frame()
            self.frames[tab_name] = new_scroll_area
            self.create_dynamic_tab(new_scroll_area, headers, tab_name)
            self.layout.addWidget(new_scroll_area)
            self.dropdown.setCurrentIndex(self.dropdown.count() - 1)
            self.show_frame(self.dropdown.count() - 1)

            # Clear the headers list and input fields
            self.listbox_headers.clear()
            self.header_input.clear()
            self.tab_name_input.clear()

            # Save parameters to file
            self.save_parameters()

    def create_dynamic_tab(self, scroll_area, headers, tab_name):
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        entries = {}
        for header in headers:
            h_layout = QHBoxLayout()
            label = QLabel(header)
            entry = QLineEdit()
            h_layout.addWidget(label)
            h_layout.addWidget(entry)
            layout.addLayout(h_layout)
            entries[header] = entry

        save_button = QPushButton("Sauvegarder")
        save_button.clicked.connect(lambda: self.save_dynamic_data(entries, tab_name))
        layout.addWidget(save_button)

        scroll_area.setWidget(scroll_content)

    def save_dynamic_data(self, entries, tab_name):
        data = {header: [entry.text()] for header, entry in entries.items()}
        filename = f'D:/Projet stage 2e/{tab_name}.xlsx'
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        QMessageBox.information(self, "Succès", "Le fichier Excel a été créé avec succès.")

    def create_scrollable_frame(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        return scroll_area

    def show_frame(self, index):
        for frame in self.frames.values():
            frame.setVisible(False)
        frame = self.frames[self.dropdown.itemText(index)]
        frame.setVisible(True)

    def save_parameters(self):
        parameters = {
            "tabs": []
        }
        for i in range(1, self.dropdown.count()):
            tab_name = self.dropdown.itemText(i)
            headers = []
            layout = self.frames[tab_name].widget().layout()
            for j in range(layout.count() - 1):  # Exclude the last item (save button)
                item = layout.itemAt(j)
                if isinstance(item, QHBoxLayout):
                    label = item.itemAt(0).widget()
                    if isinstance(label, QLabel):
                        headers.append(label.text())
            parameters["tabs"].append({"tab_name": tab_name, "headers": headers})
        with open('D:/Projet stage 2e/Projet_excel/parameters.json', 'w') as file:
            json.dump(parameters, file)

    def load_parameters(self):
        try:
            with open('D:/Projet stage 2e/Projet_excel/parameters.json', 'r') as file:
                parameters = json.load(file)
                for tab in parameters["tabs"]:
                    tab_name = tab["tab_name"]
                    headers = tab["headers"]
                    self.dropdown.addItem(tab_name)
                    new_scroll_area = self.create_scrollable_frame()
                    self.frames[tab_name] = new_scroll_area
                    self.create_dynamic_tab(new_scroll_area, headers, tab_name)
                    self.layout.addWidget(new_scroll_area)
        except FileNotFoundError:
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())