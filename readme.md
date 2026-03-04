Backend SPARTA
=================
This is the backend repository for the SPARTA application, which serves as the core engine for managing various operations and data processing tasks.
It is built using Python and Flask, providing a robust API for frontend applications to interact with.
Features
--------
- RESTful API endpoints for data management
- User authentication and authorization
- Integration with external services
- Data validation and error handling
Getting Started
---------------
To get started with the SPARTA backend, follow these steps:
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/sparta-backend.git
    cd sparta-backend
    ```
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set up the database:
   ```bash
   Using Spreeadsheet or your preferred database setup method.
   ```
4. Run the development server:
   ```bash
   python app.py
   ```
5. Access the API at `https://sparta-backend.onrender.com/api`.
Configuration
-------------
Configuration settings can be found in the `config.py` file. Make sure to adjust the settings according to your environment.

Dokumentasi manual alur revisi RAB
1. Ditolak koordinator
* Ubah status nya Ditolak oleh Koordinator pada Data form (form2) berdasarkan nomor ulok dan lingkup
* Ubah statusnya menjadi Active pada spreadsheet Gantt Chart DB (gantt_chart) berdasarkan nomor ulok dan lingkup
2. Ditolak manager
* Ubah statusnya Ditolak oleh Manajer pada Data form (form2) berdasarkan nomor ulok dan lingkup
* Hapus baris data yang masuk pada form3 berdasarkan nomor ulok dan lingkup
* Ubah statusnya menjadi Active pada spreadsheet Gantt Chart DB (gantt_chart) berdasarkan nomor ulok dan lingkup
