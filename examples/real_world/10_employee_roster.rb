# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("10_employee_roster.xlsx") do |wb|
  wb.sheet("Employees") do |sheet|
    sheet.row(["ID", "Name", "Department", "Role", "Hire Date", "Salary"])
    sheet.row(["EMP-101", "Alice Smith", "Engineering", "Senior Dev", Date.new(2020, 1, 15), 120_000])
    sheet.row(["EMP-102", "Bob Jones", "Sales", "Manager", Date.new(2018, 5, 20), 95_000])
    sheet.row(["EMP-103", "Charlie Brown", "HR", "Recruiter", Date.new(2022, 11, 1), 75_000])
  end
end
