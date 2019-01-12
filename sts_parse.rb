require "roo"
require "roo-xls"
require "csv"
require "pry"
require "unicode_utils"
require "geocoder"

def get_key_pair_values(sheet, offset, column_key, column_value, initial_offset = 15)
  i = initial_offset
  values = []
  while sheet[i + offset] && sheet[i + offset][column_key]
    values << sheet[i + offset][column_key]
    values << sheet[i + offset][column_value]
    i += 1
  end
  return values
end

def get_values(sheet, offset, column_value, initial_offset = 15)
  i = initial_offset
  values = []
  while sheet[i + offset] && sheet[i + offset][column_value] &&
        !sheet[i + offset][column_value].to_s.start_with?("РАЗДЕЛ II. БЛИЗКИЕ РОДСТВЕННИКИ") &&
        !sheet[i + offset][column_value].to_s.start_with?("FORM STI - 155 - 014")
    values << sheet[i + offset][column_value]
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
  workplace = workplace.gsub(/р-н/, "").gsub(/^г./, "").gsub(/района/, "").gsub(/району/, "").gsub(/район/, "").gsub(/УГНС/, "").gsub(/\d+/, "").gsub(/\./, "").gsub(/по /, "").gsub(/р\/н/, "").gsub(/^\s+/, "").gsub(/\s+$/, "").gsub(/ /, "").gsub(/н-н/, "").strip
  unless @geo
    @geo = {}
    CSV.readlines("geo.csv").each do |l|
      @geo[l[0]] = {lat: l[1], lon: l[2], area: l[3], city: l[4]}
    end
  end
  if @geo[workplace]
    @geo[workplace]
  else
    {:lat => "42.8807207", :lon => "74.6092764", :area => "Чуйская область", :city => "Бишкек"}
  end
end

all_values = []

max_estate_length = 0
max_property_length = 0
max_income_length = 1
max_outcome_length = 1

max_relative_estate_length = 0
max_relative_property_length = 0
max_relative_income_length = 0
max_relative_outcome_length = 0

Dir.glob("files/2017/*.csv").each_with_index do |f, index|
  puts "Processing:#{f}, index: #{index}"

  workbook = CSV.readlines(f)

  year = 2017
  offsets = workbook.each_index.select { |i| workbook[i][1] && workbook[i][1] == "FORM STI - 155 - 014" }

  offsets.each do |offset|
    offset_start_row = (1..10).find { |x| workbook[offset + x][1] && workbook[offset + x][1].include?("налоговый") }
    inn = workbook[offset + offset_start_row + 1][1].strip
    offset_start_row = (1..10).find { |x| workbook[offset + x][1] && (workbook[offset + x][1].include?("Ф.И.О.") || workbook[offset + x][1].include?("ФИО") || workbook[offset + x][1].include?("Ф.И.О")) }
    workplace = workbook[offset + offset_start_row - 1][4].strip
    full_name = (workbook[offset + offset_start_row + 1][4] || workbook[offset + offset_start_row + 1][3] || workbook[offset + offset_start_row][4] || workbook[offset + offset_start_row][3]).strip
    position = (workbook[offset + offset_start_row + 1][9] || workbook[offset + offset_start_row][9]).strip
    place = find_location(workplace)
    cleaned_position = clean_position(position)
    gender = detect_gender(full_name)

    puts "f:#{f},offset:#{offset},n:#{full_name.strip},inn:#{inn},position:#{position.strip},workplace:#{workplace.strip}"

    income = get_values(workbook, offset, 1, offset_start_row + 7)
    outcome = get_values(workbook, offset, 2, offset_start_row + 7)

    real_estate = get_key_pair_values(workbook, offset, 3, 4, offset_start_row + 7)
    max_estate_length = real_estate.size if real_estate.size > max_estate_length.to_i

    property = get_key_pair_values(workbook, offset, 7, 8, offset_start_row + 7)
    max_property_length = property.size if property.size > max_property_length.to_i

    relative_start_row = (10..100).find { |x| workbook[offset + x][1] && workbook[offset + x][1].start_with?("РАЗДЕЛ II") }

    relative_income = get_values(workbook, offset, 1, relative_start_row + 1)
    max_relative_income_length = relative_income.size if relative_income.size > max_relative_income_length.to_i

    relative_outcome = get_values(workbook, offset, 2, relative_start_row + 1)
    max_relative_outcome_length = relative_outcome.size if relative_outcome.size > max_relative_outcome_length.to_i

    relative_real_estate = get_key_pair_values(workbook, offset, 3, 4, relative_start_row + 1)
    max_relative_estate_length = relative_real_estate.size if relative_real_estate.size > max_relative_estate_length.to_i

    relative_property = get_key_pair_values(workbook, offset, 7, 8, relative_start_row + 1)
    max_relative_property_length = relative_property.size if relative_property.size > max_relative_property_length.to_i

    begin
      reporter_income = income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f
      spouse_income = relative_income.map { |i| i.to_s.gsub(/,/, ".").gsub(/\s/, "").scan(/\d+\.?\d+/) }.flatten.map(&:to_f).inject(:+).to_f
      all_values << {f: f, year: year, inn: inn, workplace: workplace, position: position, cleaned_position: cleaned_position, place: place, full_name: full_name, gender: gender,
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
end

CSV.open("report_2017.csv", "w") do |csv|
  csv << ["File", "Year", "INN", "Workplace", "Position", "CleanedPosition", "PlaceLat", "PlaceLon", "PlaceArea", "PlaceCity", "FullName", "Gender", "TotalIncome", "ReporterIncome", "SpouseIncome",
          (1..max_income_length).map { |i| "Income#{i}" },
          (1..max_outcome_length).map { |i| "Outcome#{i}" },
          (1..max_estate_length / 2).map { |i| ["RealEstateType#{i}", "RealEstateArea#{i}"] },
          (1..max_property_length / 2).map { |i| ["PropertyType#{i}", "PropertyDescription#{i}"] },
          (1..max_relative_income_length).map { |i| "RelativeIncome#{i}" },
          (1..max_relative_outcome_length).map { |i| "RelativeOutcome#{i}" },
          (1..max_relative_estate_length / 2).map { |i| ["RelativeRealEstateType#{i}", "RelativeRealEstateArea#{i}"] },
          (1..max_relative_property_length / 2).map { |i| ["RelativePropertyType#{i}", "RelativePropertyDescription#{i}"] }].flatten
  all_values.each do |v|
    csv << [v[:f], v[:year], v[:inn], v[:workplace], v[:position], v[:cleaned_position], v[:place][:lat], v[:place][:lon], v[:place][:area], v[:place][:city],
            v[:full_name], v[:gender], v[:total_income], v[:reporter_income], v[:spouse_income],
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
