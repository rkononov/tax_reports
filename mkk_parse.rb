require "rubyXL"
require "csv"
require "pry"

def get_key_pair_values(sheet, offset, column_key, column_value, initial_offset = 14)
  i = initial_offset
  values = []
  while sheet[i - offset] && sheet[i - offset][column_key] && sheet[i - offset][column_key].value
    values << sheet[i - offset][column_key].value
    values << sheet[i - offset][column_value].value
    i += 1
  end
  return values
end

def get_values(sheet, offset, column_value, initial_offset = 14)
  i = initial_offset
  values = []
  while sheet[i - offset] && sheet[i - offset][column_value] && sheet[i - offset][column_value].value &&
        sheet[i - offset][column_value].value != "Близкий родственник" &&
        !sheet[i - offset][column_value].value.to_s.start_with?("* должность указана")
    values << sheet[i - offset][column_value].value
    i += 1
  end
  return values
end

def format_values(values, max)
  values.fill("", values.size..max - 1)
end

all_values = []

max_estate_length = 0
max_property_length = 0
max_income_length = 0
max_outcome_length = 0

max_relative_estate_length = 0
max_relative_property_length = 0
max_relative_income_length = 0
max_relative_outcome_length = 0

Dir.glob("files/*.xlsx").each_with_index do |f, index|
  puts "Processing:#{f}, index: #{index}"
  workbook = RubyXL::Parser.parse(f)
  sheet = workbook[0]
  year = f[6..9].to_i - 1
  offset = 4 - (1..10).find { |x| sheet[x][0].value && sheet[x][0].value[0..11] == "Наименование" }
  workplace = sheet[4 - offset][7].value
  full_name = sheet[7 - offset][7].value
  position = sheet[8 - offset][7].value

  income = get_values(sheet, offset, 0)
  max_income_length = income.size if income.size > max_income_length

  outcome = get_values(sheet, offset, 2)
  max_outcome_length = outcome.size if outcome.size > max_outcome_length

  real_estate = get_key_pair_values(sheet, offset, 4, 7)
  max_estate_length = real_estate.size if real_estate.size > max_estate_length

  property = get_key_pair_values(sheet, offset, 13, 15)
  max_property_length = property.size if property.size > max_property_length

  relative_start_row = (10..100).find { |x| sheet[x] && sheet[x][0] && sheet[x][0].value == "Близкий родственник" }

  relative_income = get_values(sheet, offset, 0, relative_start_row + 1 + offset)
  max_relative_income_length = relative_income.size if relative_income.size > max_relative_income_length

  relative_outcome = get_values(sheet, offset, 2, relative_start_row + 1 + offset)
  max_relative_outcome_length = relative_outcome.size if relative_outcome.size > max_relative_outcome_length

  relative_real_estate = get_key_pair_values(sheet, offset, 4, 7, relative_start_row + 1 + offset)
  max_relative_estate_length = relative_real_estate.size if relative_real_estate.size > max_relative_estate_length

  relative_property = get_key_pair_values(sheet, offset, 13, 15, relative_start_row + 1 + offset)
  max_relative_property_length = relative_property.size if relative_property.size > max_relative_property_length
  begin
    all_values << {f: f, year: year, workplace: workplace, position: position, full_name: full_name,
                   total_income: income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f +
                                 relative_income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f,
                   income: income, outcome: outcome,
                   real_estate: real_estate, property: property,
                   relative_income: relative_income, relative_outcome: relative_outcome,
                   relative_real_estate: relative_real_estate, relative_property: relative_property}
  rescue => ex
    puts "EX:#{ex.inspect}"
    puts "income:#{income}"
    puts "relative_income:#{relative_income}"
    raise ex
  end
end

# income.gsub(/,/,'.').gsub(/\s/,'').scan(/\d+\.?\d+/)
# all_values.map{|x|x[:f] if x[:income].any?{|v|v.to_s.include?("$")||v.to_s.include?("долл")}}.uniq

CSV.open("report.csv", "w") do |csv|
  csv << ["File", "Year", "Workplace", "Position", "FullName", "TotalIncome",
          (1..max_income_length).map { |i| "Income#{i}" },
          (1..max_outcome_length).map { |i| "Outcome#{i}" },
          (1..max_estate_length / 2).map { |i| ["RealEstateType#{i}", "RealEstateArea#{i}"] },
          (1..max_property_length / 2).map { |i| ["PropertyType#{i}", "PropertyDescription#{i}"] },
          (1..max_relative_income_length).map { |i| "RelativeIncome#{i}" },
          (1..max_relative_outcome_length).map { |i| "RelativeOutcome#{i}" },
          (1..max_relative_estate_length / 2).map { |i| ["RelativeRealEstateType#{i}", "RelativeRealEstateArea#{i}"] },
          (1..max_relative_property_length / 2).map { |i| ["RelativePropertyType#{i}", "RelativePropertyDescription#{i}"] }].flatten
  all_values.each do |v|
    csv << [v[:f], v[:year], v[:workplace], v[:position], v[:full_name], v[:total_income],
            format_values(v[:income], max_income_length),
            format_values(v[:outcome], max_outcome_length),
            format_values(v[:real_estate], max_estate_length),
            format_values(v[:property], max_property_length),
            format_values(v[:relative_income], max_relative_income_length),
            format_values(v[:relative_outcome], max_relative_outcome_length),
            format_values(v[:relative_real_estate], max_relative_estate_length),
            format_values(v[:relative_property], max_relative_property_length)].flatten
  end
end
pry
