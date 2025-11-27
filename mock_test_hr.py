import unittest
from unittest.mock import MagicMock, patch
from datetime import date, time, datetime
import sys

# Mock psycopg2 before importing hr_database
sys.modules['psycopg2'] = MagicMock()
sys.modules['psycopg2.extras'] = MagicMock()

# Now import the module to test
import hr_database

class TestProcessAttendance(unittest.TestCase):
    @patch('hr_database.get_db_connection')
    def test_logic(self, mock_get_conn):
        # Setup Mock DB
        mock_conn = MagicMock()
        mock_cursor = MagicMock()
        mock_get_conn.return_value = mock_conn
        mock_conn.cursor.return_value.__enter__.return_value = mock_cursor

        # Mock Data
        # 1. Employees
        mock_cursor.fetchall.side_effect = [
            [ # Employees
                {'emp_id': 'EMP001', 'fname': 'John', 'lname': 'Doe', 'work_location': 'สำนักงานใหญ่'},
                {'emp_id': 'EMP002', 'fname': 'Jane', 'lname': 'Smith', 'work_location': 'คลังสินค้า'}
            ],
            [ # Holidays
                # No holidays
            ],
            [ # Leaves
                # No leaves
            ],
            [ # Logs
                {'emp_id': 'EMP001', 'scan_timestamp': datetime(2023, 11, 1, 8, 55, 0)}, # On time
                {'emp_id': 'EMP001', 'scan_timestamp': datetime(2023, 11, 1, 18, 5, 0)}, # OT 5 min
                {'emp_id': 'EMP002', 'scan_timestamp': datetime(2023, 11, 1, 8, 10, 0)}, # Late (Warehouse starts 8:00)
                {'emp_id': 'EMP002', 'scan_timestamp': datetime(2023, 11, 1, 17, 0, 0)}, # On time out
            ]
        ]

        # Run function
        start_date = date(2023, 11, 1)
        end_date = date(2023, 11, 1)
        results = hr_database.process_attendance_summary(start_date, end_date)

        # Verify
        print(f"Results: {results}")
        self.assertEqual(len(results), 2)
        
        # Check EMP001 (HQ: 9:00-18:00)
        emp1 = next(r for r in results if r['emp_id'] == 'EMP001')
        self.assertEqual(emp1['status'], 'Normal')
        self.assertEqual(emp1['ot_minutes'], 5)

        # Check EMP002 (Warehouse: 8:00-17:00)
        emp2 = next(r for r in results if r['emp_id'] == 'EMP002')
        self.assertEqual(emp2['status'], 'Late (<30m)') # 8:10 is late for 8:00
        self.assertEqual(emp2['late_minutes'], 60) # Penalty 1

if __name__ == '__main__':
    unittest.main()
