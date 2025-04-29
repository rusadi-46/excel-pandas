# Pada file yang menggunakan class SalarySlipManager
from salary_slip_manager import SalarySlipManager

class DummySalary:
    @staticmethod
    def insert_salary_data():
        manager = SalarySlipManager()
        
        staff_data = {
            'full_name': 'Staff 1',
            'alias': 'staff 1',
            'gender': 'female',
            'email': 'staff@gmail.com',
            'employee_number': '123124124124',
            'mobile_number': '123124124124',
            'position': 'Stylist',
            'identity_number': '3267123123123123',
            'join_date': '2022-06-01'
        }

        salary_data = {
            'period_month': '2025-04',
            'total_income': 5087425,
            'total_deduction': 15000,
            'net_salary': 5072425
        }

        income_items = [
            ('Gaji Pokok', 1500000),
            ('Komisi', 3497925),
            ('Lembur', 30000),
            ('BPJS', 59500)
        ]

        deduction_items = [
            ('Telat', 15000)
        ]

        attendance_data = {
            'hadir': 22,
            'sakit': 2,
            'izin': 1,
            'cuti': 0,
            'lembur': 2,
            'telat': 1
        }

        try:
            staff_id = manager.insert_staff(
                staff_data['full_name'],
                staff_data['alias'],
                staff_data['gender'],
                staff_data['email'],
                staff_data['employee_number'],
                staff_data['mobile_number'],
                staff_data['position'],
                staff_data['identity_number'],
                staff_data['join_date'],
                staff_data.get('status', 'Active')
            )

            slip_id = manager.insert_salary_slip(
                staff_id=staff_id,
                period_month=salary_data['period_month'],
                total_income=salary_data['total_income'],
                total_deduction=salary_data['total_deduction'],
                net_salary=salary_data['net_salary']
            )

            manager.insert_salary_items(
                slip_id,
                income_items=income_items,
                deduction_items=deduction_items
            )

            manager.insert_attendance_record(
                slip_id,
                hadir=attendance_data.get('hadir', 0),
                sakit=attendance_data.get('sakit', 0),
                izin=attendance_data.get('izin', 0),
                cuti=attendance_data.get('cuti', 0),
                lembur=attendance_data.get('lembur', 0),
                telat=attendance_data.get('telat', 0)
            )

            manager.commit()
            print("✅ Data gaji berhasil disimpan.")
        except Exception as e:
            manager.rollback()
            print(f"❌ Terjadi error: {e}")
        finally:
            manager.close()


# Contoh Pemanggilan
if __name__ == "__main__":
    DummySalary.insert_salary_data()
