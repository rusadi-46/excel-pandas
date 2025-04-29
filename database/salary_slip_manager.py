from .db_connection import create_connection
import mysql.connector
from mysql.connector import Error

class SalarySlipManager:
    def __init__(self):
        self.conn = create_connection()
        if self.conn:
            self.cursor = self.conn.cursor(dictionary=True)
        else:
            self.cursor = None
            print("❌ Gagal membuat koneksi.")

    def close(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def insert_staff(self, full_name, alias, gender, email, employee_number, mobile_number, position, identity_number, join_date, status='Active'):
        sql = """
            INSERT INTO staff (full_name, alias, gender, email, employee_number, mobile_number, position, identity_number, join_date, status) 
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        self.cursor.execute(sql, (full_name, alias, gender, email, employee_number, mobile_number, position, identity_number, join_date, status))
        return self.cursor.lastrowid

    def insert_salary_slip(self, staff_id, period_month, total_income, total_deduction, net_salary):
        sql = """
            INSERT INTO salary_slip (staff_id, period_month, total_income, total_deduction, net_salary)
            VALUES (%s, %s, %s, %s, %s)
        """
        self.cursor.execute(sql, (staff_id, period_month, total_income, total_deduction, net_salary))
        return self.cursor.lastrowid

    def insert_salary_items(self, slip_id, income_items=[], deduction_items=[]):
        sql = """
            INSERT INTO salary_item (slip_id, item_type, item_name, amount) 
            VALUES (%s, %s, %s, %s)
        """
        for item_name, amount in income_items:
            self.cursor.execute(sql, (slip_id, 'income', item_name, amount))
        
        for item_name, amount in deduction_items:
            self.cursor.execute(sql, (slip_id, 'deduction', item_name, amount))

    def insert_attendance_record(self, slip_id, hadir=0, sakit=0, izin=0, cuti=0, lembur=0, telat=0):
        sql = """
            INSERT INTO attendance_record (slip_id, hadir, sakit, izin, cuti, lembur, telat)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        self.cursor.execute(sql, (slip_id, hadir, sakit, izin, cuti, lembur, telat))

    def fetch_salary_slip_by_alias_and_period(self, alias, period):
        try:
            query = """
                SELECT 
                    s.id AS staff_id,
                    s.name AS staff_name,
                    ss.id AS slip_id,
                    ss.period_month,
                    si.id AS salary_item_id,
                    si.item_name,
                    si.item_type,
                    si.amount,
                    ar.id AS attendance_id,
                    ar.hadir,
                    ar.sakit,
                    ar.izin,
                    ar.cuti,
                    ar.lembur,
                    ar.telat
                FROM staff s
                JOIN salary_slip ss ON ss.staff_id = s.id
                LEFT JOIN salary_item si ON si.slip_id = ss.id
                LEFT JOIN attendance_record ar ON ar.slip_id = ss.id
                WHERE s.alias = %s AND ss.period_month = %s
            """
            self.cursor.execute(query, (alias, period))
            return self.cursor.fetchall()
        except Error as e:
            print(f"❌ Query gagal: {e}")
            return []

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()
