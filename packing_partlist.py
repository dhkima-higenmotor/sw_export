import tkinter as tk
from tkinter import filedialog
import os
import shutil
import pandas as pd

class FilePackerApp:
    def __init__(self, master):
        self.master = master
        master.title("Copy drawing files by csv partlist")

        self.csv_path = ""
        self.work_path = ""
        self.csv_filename = ""
        self.pdf_path = []
        self.dxf_path = []
        self.step_path = []
        self.output_path = ""
        self.PartName = []
        self.pdf_path2 = []
        self.dxf_path2 = []
        self.step_path2 = []

        # CSV File Entry and Browse Button
        self.csv_label = tk.Label(master, text="CSV File:")
        self.csv_label.grid(row=0, column=0, padx=10, pady=5)

        self.csv_entry = tk.Entry(master, width=50)
        self.csv_entry.grid(row=0, column=1, padx=10, pady=5)

        self.browse_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=10, pady=5)

        self.csv_label = tk.Label(master, text="csv 파트리스트 파일을 선택하세요.")
        self.csv_label.grid(row=1, column=1, padx=10, pady=5)

        # Packing and Exit Buttons
        self.packing_button = tk.Button(master, text="PACKING", command=self.pack_files)
        self.packing_button.grid(row=2, column=0, padx=10, pady=10)

        self.exit_button = tk.Button(master, text="EXIT", command=master.quit)
        self.exit_button.grid(row=2, column=2, padx=10, pady=5)

    def browse_file(self):
        """Opens a file dialog to select a CSV file and inserts the path into the entry."""
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.csv_entry.delete(0, tk.END)
            self.csv_entry.insert(0, file_path)
        self.csv_label.config(text="PACKING 버튼을 누르면 도면 파일 복사가 됩니다다.")

    def read_csv(self):
        # Read the CSV file into a pandas DataFrame
        print("csv_path : ")
        print(self.csv_path)
        try:
            df = pd.read_csv(self.csv_path, encoding='euc-kr')
        except FileNotFoundError:
            print(f"Error: File not found at path: {self.csv_path}")
            exit()
        except pd.errors.EmptyDataError:
            print(f"Error: The file at {self.csv_path} is empty.")
            exit()
        except pd.errors.ParserError:
            print(f"Error: Could not parse the CSV file at {self.csv_path}. Check the file format.")
            exit()
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            exit()
        # Extract the "PartName" column into a list
        self.PartName = df["PartName"].tolist()
        # Print the list (optional)
        print(self.PartName)

    def pack_files(self):
        """Performs the file packing operation."""
        self.csv_path = self.csv_entry.get()

        if not self.csv_path:
            print("Error: CSV file path is empty.")
            return

        self.read_csv()

        # Split path into directory and filename
        self.work_path, self.csv_filename = os.path.split(self.csv_path)

        # Remove .csv extension
        self.csv_filename = os.path.splitext(self.csv_filename)[0]

        # Find PDF, DXF, and STEP files
        self.pdf_path = []
        self.dxf_path = []
        self.step_path = []
        self.find_files(self.work_path)

        # Create output directory
        self.output_path = os.path.join(self.work_path, self.csv_filename)
        os.makedirs(self.output_path, exist_ok=True)

        # Copy files to output directory
        self.copy_files()
        self.pdf_path = []
        self.dxf_path = []
        self.step_path = []
        self.PartName = []
        self.pdf_path2 = []
        self.dxf_path2 = []
        self.step_path2 = []
        print("Packing completed!")
        self.csv_label.config(text="파일 복사 완료!")

    def find_files(self, directory):
        """Recursively searches for PDF, DXF, and STEP files in the given directory."""
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                if file.lower().endswith(".pdf"):
                    self.pdf_path.append(file_path)
                elif file.lower().endswith(".dxf"):
                    self.dxf_path.append(file_path)
                elif file.lower().endswith(".step"):
                    self.step_path.append(file_path)

    def copy_files(self):
        """Copies the found files to the output directory."""
        for path in self.pdf_path:
            filename = os.path.basename(path)  # 파일 경로에서 파일 이름만 추출
            name_without_extension = os.path.splitext(filename)[0]  # 확장자 제거
            if name_without_extension in self.PartName:
                self.pdf_path2.append(path)
        for path in self.dxf_path:
            filename = os.path.basename(path)  # 파일 경로에서 파일 이름만 추출
            name_without_extension = os.path.splitext(filename)[0]  # 확장자 제거
            if name_without_extension in self.PartName:
                self.dxf_path2.append(path)
        for path in self.step_path:
            filename = os.path.basename(path)  # 파일 경로에서 파일 이름만 추출
            name_without_extension = os.path.splitext(filename)[0]  # 확장자 제거
            if name_without_extension in self.PartName:
                self.step_path2.append(path)
        for file_path in self.pdf_path2 + self.dxf_path2 + self.step_path2:  # 파일 카피
            try:
                shutil.copy2(file_path, self.output_path)
            except Exception as e:
                print(f"Error copying {file_path}: {e}")

root = tk.Tk()
app = FilePackerApp(root)
root.mainloop()
