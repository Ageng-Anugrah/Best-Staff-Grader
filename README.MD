# Best Staff Grader

## Overview

Aplikasi ini dibuat untuk menghitung nilai best staff dari hasil voting gform sebelum dapat menggunakan website best staff  

## Cara Penggunaan
* Pastikan file xslx yang berisi Pengurus BEM Fasilkom UI 2021 dan hasil voting dari gform berada pada folder yang sama dengan program ini. Jika tidak berada folder yang sama bisa menggunakan absolute path.

* Ubah Line 22 pada program menjadi:
    ```bash
    df = pd.read_excel (r'<nama file pengurus BEM.xslx>')
    ```
    * contoh:
        * `df = pd.read_excel (r'Pengurus BEM Fasilkom UI 2021.xlsx')`

* Ubah Line 44 pada program menjadi:
    ```bash
    sheet_to_df_map = pd.read_excel('<nama file hasil voting best staff.xslx>', sheet_name=None)
    ```
    * contoh:
        * `sheet_to_df_map = pd.read_excel('Data Best Staff April.xlsx', sheet_name=None)`

* Program akan menghasilkan 2 file output yaitu:
    * `Detail Perbirdep.xslx` yang berisikan detail nilai untuk masing-masing birdep di setiap sheetnya
    * `Nilai Akhir.xlsx` yang berisikan nilai seluruh staff BEM Fasilkom UI terurut berdasarkan nilai akhir (dari yang terbesar ke terkecil) 

## Acknowledgements

* **Developer** : PTI BEM Fasilkom UI 2021
* **Product Owner** : PSDM BEM Fasilkom UI 2021
