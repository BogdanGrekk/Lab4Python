import csv
from faker import Faker
import random
from datetime import datetime

class EmployeeDataGenerator:

    def __init__(self, file_path, record_count=2000):
        self.file_path = file_path
        self.record_count = record_count
        self.fake_data_generator = Faker()

    def generate_employee_record(self):
        """Generates a random employee record with a Ukrainian-style patronymic in English context."""
        gender = 'male' if random.random() < 0.6 else 'female'

        if gender == 'male':
            first_name = self.fake_data_generator.first_name_male()
            last_name = self.fake_data_generator.last_name_male()
            father_name = self.fake_data_generator.first_name_male()
            middle_name = father_name + 'ovich'
        else:
            first_name = self.fake_data_generator.first_name_female()
            last_name = self.fake_data_generator.last_name_female()
            father_name = self.fake_data_generator.first_name_male()
            middle_name = father_name + 'ivna'

        dob = self.fake_data_generator.date_of_birth(tzinfo=None, minimum_age=15, maximum_age=85)
        job = self.fake_data_generator.job()
        city = self.fake_data_generator.city()
        address = self.fake_data_generator.address().replace('\n', ', ')
        phone = self.fake_data_generator.phone_number()
        email = self.fake_data_generator.email()

        return [last_name, first_name, middle_name, gender, dob, job, city, address, phone, email]

    def generate_csv_file(self):
        """Generates a CSV file with employee records and includes a BOM."""
        with open(self.file_path, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerow(["Прізвище", "Ім’я", "По батькові", "Стать", "Дата народження",
                             "Посада", "Місто проживання", "Адреса проживання", "Телефон", "Email"])
            for _ in range(self.record_count):
                writer.writerow(self.generate_employee_record())

    def run(self):
        self.generate_csv_file()
        print(f"Generated {self.record_count} records in {self.file_path}")


employee_data_generator = EmployeeDataGenerator("employees.csv")
employee_data_generator.run()
