import sqlite3
import os

class DBManager:
    def __init__(self, db_path='data/db/app_database.db'):
        self.db_path = db_path
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
        # Initialize by creating a default table if not exists
        self._initialize_db()

    def _initialize_db(self):
        # Example: Create a simple table for processed PDF files metadata
        # This is just an example; actual tables will depend on application needs.
        # For now, the PDF processor outputs to Excel, so DB use for PDFs is minimal.
        # This table could store, e.g., PDF filename, processing date, output Excel path.
        pdf_metadata_table_sql = """
        CREATE TABLE IF NOT EXISTS pdf_processing_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pdf_filename TEXT NOT NULL,
            processed_at TEXT NOT NULL,
            output_excel_path TEXT,
            status TEXT,
            notes TEXT
        );
        """
        try:
            self.create_table(pdf_metadata_table_sql)
            print(f"Database initialized/checked at {self.db_path}")
        except sqlite3.Error as e:
            print(f"Error initializing database table: {e}")


    def connect(self):
        """Creates and returns a database connection."""
        try:
            conn = sqlite3.connect(self.db_path)
            return conn
        except sqlite3.Error as e:
            print(f"Error connecting to database {self.db_path}: {e}")
            return None

    def close(self, conn):
        """Closes the given database connection."""
        if conn:
            conn.close()

    def execute_query(self, query, params=()):
        """Executes a given SQL query (e.g., INSERT, UPDATE, DELETE)."""
        conn = self.connect()
        if not conn:
            return False
        try:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error executing query: {e}")
            return False
        finally:
            self.close(conn)
            
    def fetch_query(self, query, params=()):
        """Executes a SELECT query and returns fetched results."""
        conn = self.connect()
        if not conn:
            return None
        try:
            cursor = conn.cursor()
            cursor.execute(query, params)
            results = cursor.fetchall()
            return results
        except sqlite3.Error as e:
            print(f"Error fetching query: {e}")
            return None
        finally:
            self.close(conn)

    def create_table(self, table_definition_sql):
        """Creates a table based on the SQL definition."""
        return self.execute_query(table_definition_sql)

if __name__ == '__main__':
    # Example Usage
    # This example assumes the script is run from the project root directory
    
    # Create a DBManager instance. If 'data/db/app_database.db' or its directory doesn't exist, they will be created.
    db_manager = DBManager(db_path='data/db/app_database.db') # Initializes pdf_processing_log table

    # Example: Log a processed PDF (hypothetical)
    from datetime import datetime
    current_time = datetime.now().isoformat()
    
    log_success = db_manager.execute_query(
        "INSERT INTO pdf_processing_log (pdf_filename, processed_at, output_excel_path, status, notes) VALUES (?, ?, ?, ?, ?)",
        ("example.pdf", current_time, "data/exports/example_improved.xlsx", "success", "Test entry")
    )
    if log_success:
        print("Successfully logged a dummy PDF processing entry.")
    else:
        print("Failed to log dummy PDF processing entry.")

    # Example: Fetch all logs
    all_logs = db_manager.fetch_query("SELECT * FROM pdf_processing_log")
    if all_logs is not None:
        print(f"Current PDF processing logs (count: {len(all_logs)}):")
        for row in all_logs:
            print(row)
    else:
        print("Failed to fetch logs.")

    # Example: Create another table for QC data (as per Markdown's measurement.db)
    # This would typically be more complex and potentially in a separate DB or managed differently.
    qc_measurements_table_sql = """
    CREATE TABLE IF NOT EXISTS qc_measurements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        part_name TEXT,
        serial_number TEXT,
        feature_name TEXT,
        nominal REAL,
        plus_tol REAL,
        minus_tol REAL,
        meas REAL,
        dev REAL,
        outtol REAL,
        timestamp TEXT
    );
    """
    if db_manager.create_table(qc_measurements_table_sql):
        print("qc_measurements table created or already exists.")
        
        # Add some dummy QC data
        qc_data_success = db_manager.execute_query(
            "INSERT INTO qc_measurements (part_name, serial_number, feature_name, nominal, timestamp) VALUES (?, ?, ?, ?, ?)",
            ("PartX", "SN001", "DimA", 50.0, current_time)
        )
        if qc_data_success:
            print("Dummy QC data added.")
        
        qc_entries = db_manager.fetch_query("SELECT * FROM qc_measurements")
        if qc_entries:
            print(f"QC Entries (count: {len(qc_entries)}):")
            for entry in qc_entries:
                print(entry)
    else:
        print("Failed to create qc_measurements table.")
