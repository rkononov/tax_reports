require "roo"
require "roo-xls"
require "csv"
require "pry"
require "unicode_utils"

def get_key_pair_values(sheet, offset, column_key, column_value, initial_offset = 15)
  i = initial_offset
  values = []
  while sheet.cell(i - offset, column_key)
    values << sheet.cell(i - offset, column_key)
    values << sheet.cell(i - offset, column_value)
    i += 1
  end
  return values
end

def get_values(sheet, offset, column_value, initial_offset = 15)
  i = initial_offset
  values = []
  while sheet.cell(i - offset, column_value) &&
        sheet.cell(i - offset, column_value) != "Близкий родственник" &&
        !sheet.cell(i - offset, column_value).to_s.start_with?("* должность указана")
    values << sheet.cell(i - offset, column_value)
    i += 1
  end
  return values
end

def format_values(values, max)
  values.fill("", values.size..max - 1)
end

def detect_gender(full_name)
  name = full_name.split(" ")
  return "Ж" if name[2] && name[2].end_with?("вна")
  return "Ж" if name.count == 2 && name[0].end_with?("а")
  return "М"
end

def clean_position(pos)
  return pos unless pos
  d = UnicodeUtils.downcase(pos.strip.lstrip)
  d.gsub!(/экс-/, "")
  %w(посол консул эксперт депутат директор судья).map do |p|
    return p if d.include?(p)
  end
  return "таможенник" if d.include?("таможн")
  return "председатель" if d == "и.о председателя" || d.include?("председатель")
  return "вице-мэр" if d == "<html>вице-мэр<b>     </b></html>" || d == "<html>вице-мэр<b>     </b></html>"
  return "заведующий сектором" if d == "заведующая сектором"
  return "заместитель начальника" if d.start_with?("заместитель начальника")
  return "начальник угнс" if d.start_with?("начальник угнс")
  return d
end

def find_location(workplace)
  return workplace unless workplace
  d = UnicodeUtils.downcase(workplace.strip.lstrip)
  return "кызыл-кыя" if d.include?("кызыл-кыя") || d.include?("кызыл-ки")
  %w(первомай каракуль кочкор сузак араван ак-суй базар-коргон токтогул ноокат орловка кара-кол кок-жангак алай
     кемин манас каинда кочкор-ата ак-тал ат-баш исфана тон ноокен панфилов кант кербен шопоков кадамжай
     жайыл аламудун сокулук каракол тогуз-торо джумгал московский чон-алай тюп жайыл бишкек ош майлуу-суу баткен кара-балта
     талас жалал-абад нарын лейлек чолпон-ата сулюкта аксый узген аксуй кара-куль токмок айдаркен чаткал чуй иссык-куль).map do |city|
    return city if d.include?(city)
  end
  return "джети-огуз" if d.include?("джети-ог") || d.include?("джети ог") || d.include?("джеты-ог")
  return "таш-комур" if d.include?("таш-к")
  return "кара-буура" if d.include?("кара-бу")
  return "кара-суу" if d.include?("кара-су")
  return "балыкчи" if d.include?("балыкч")
  return "чолпон-ата" if d.include?("чолпон ата")
  return "ала-бука" if d.include?("ала-бук")
  return "московский" if d.include?("московская")
  return "ысык-ата" if d.include?("ысык-ат")
  return "кара кулжа" if d.include?("кара кулж") || d.include?("кара-кулж")
  return "сулюкта" if d.include?("сулюкт")
  return "бакай ата" if d.include?("бакай-ат")
  return "аламудун" if d.include?("аламунская")
  return "майлуу-суу" if d.include?("майлуу-су")
  return "талас" if d.include?("тааласский")
  return "бишкек" if %w(полномочное агентстве верзовный миграции агентсво институт правительство военный предствительство центр кенеш комитет гвардия комитет фон национальный штаб министер агентство служба департамент аппарат инспекц комиссия палата прокуратура верховный управление).any? { |v| d.include?(v) }
  return d
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
  next if f[0..2] == "bad"

  workbook = begin
               Roo::Spreadsheet.open(f)
             rescue
               Roo::Spreadsheet.open(f, extension: :xls)
             end
  sheet = (workbook.sheets.count > 3) ? workbook.sheet(workbook.sheets.count - 1) : workbook.sheet(0)
  year = f[6..9].to_i - 1
  offset = 5 - (1..10).find { |x| sheet.cell(x, 1) && sheet.cell(x, 1)[0..11] == "Наименование" }
  workplace = sheet.cell(5 - offset, 8)
  place = find_location(workplace)
  full_name = sheet.cell(8 - offset, 8)
  position = sheet.cell(9 - offset, 8)
  cleaned_position = clean_position(position)
  gender = detect_gender(full_name)

  income = get_values(sheet, offset, 1)
  max_income_length = income.size if income.size > max_income_length

  outcome = get_values(sheet, offset, 3)
  max_outcome_length = outcome.size if outcome.size > max_outcome_length

  real_estate = get_key_pair_values(sheet, offset, 5, 8)
  max_estate_length = real_estate.size if real_estate.size > max_estate_length

  property = get_key_pair_values(sheet, offset, 14, 16)
  max_property_length = property.size if property.size > max_property_length

  relative_start_row = (10..100).find { |x| sheet.cell(x, 1) && sheet.cell(x, 1) == "Близкий родственник" }

  relative_income = get_values(sheet, offset, 1, relative_start_row + 1 + offset)
  max_relative_income_length = relative_income.size if relative_income.size > max_relative_income_length

  relative_outcome = get_values(sheet, offset, 3, relative_start_row + 1 + offset)
  max_relative_outcome_length = relative_outcome.size if relative_outcome.size > max_relative_outcome_length

  relative_real_estate = get_key_pair_values(sheet, offset, 5, 8, relative_start_row + 1 + offset)
  max_relative_estate_length = relative_real_estate.size if relative_real_estate.size > max_relative_estate_length

  relative_property = get_key_pair_values(sheet, offset, 14, 16, relative_start_row + 1 + offset)
  max_relative_property_length = relative_property.size if relative_property.size > max_relative_property_length
  begin
    reporter_income = income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f
    spouse_income = relative_income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f
    all_values << {f: f, year: year, workplace: workplace, position: position, cleaned_position: cleaned_position, place: place, full_name: full_name, gender: gender,
                   total_income: reporter_income + spouse_income,
                   reporter_income: reporter_income, spouse_income: spouse_income,
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

CSV.open("report.csv", "w") do |csv|
  csv << ["File", "Year", "Workplace", "Position", "CleanedPosition", "Place", "FullName", "Gender", "TotalIncome", "ReporterIncome", "SpouseIncome",
          (1..max_income_length).map { |i| "Income#{i}" },
          (1..max_outcome_length).map { |i| "Outcome#{i}" },
          (1..max_estate_length / 2).map { |i| ["RealEstateType#{i}", "RealEstateArea#{i}"] },
          (1..max_property_length / 2).map { |i| ["PropertyType#{i}", "PropertyDescription#{i}"] },
          (1..max_relative_income_length).map { |i| "RelativeIncome#{i}" },
          (1..max_relative_outcome_length).map { |i| "RelativeOutcome#{i}" },
          (1..max_relative_estate_length / 2).map { |i| ["RelativeRealEstateType#{i}", "RelativeRealEstateArea#{i}"] },
          (1..max_relative_property_length / 2).map { |i| ["RelativePropertyType#{i}", "RelativePropertyDescription#{i}"] }].flatten
  all_values.each do |v|
    csv << [v[:f], v[:year], v[:workplace], v[:position], v[:cleaned_position], v[:place], v[:full_name], v[:gender], v[:total_income], v[:reporter_income], v[:spouse_income],
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
