require 'simple_xlsx_reader'
require 'csv'
require 'json'
require 'pry'

class Convert
  class << self

    def to_json(file)
      xlsx = file_content_to_csv(file)
      csv_files_to_json(xlsx.sheets.map(&:name))
    end

    def to_csv(file)
      file_content_to_csv(file)
    end

    private
    def file_content_to_csv(file)

      xlsx = SimpleXlsxReader.open(file)

      xlsx.sheets.each do |sheet|
        file_name = build_file_name(sheet.name)
        File.open(file_name,  'w') do |f|
          sheet.rows.each do |row|
            f.write "#{row.join(',')}\n"
          end
        end
      end
      xlsx

    end

    def build_file_name(string, format=".csv")
      "#{string.gsub(/\s+/, '_').downcase}#{format}"
    end

    def csv_files_to_json(sheets_names)
      sheets_names.each do |sheet_name|
        associated_data = build_hash_object(sheet_name)
        json_content    = build_json_data(associated_data)
        File.open(build_file_name(sheet_name, '.json'), 'w') do |csv|
          csv.write json_content
        end
        #delete the csv converted to json
        File.delete(build_file_name(sheet_name))
      end
    end

    def build_json_data(data)
      JSON.pretty_generate(data)
    end

    def build_hash_object(sheet_name)
      CSV.foreach(build_file_name(sheet_name), headers: true).map do |row|
        [ row[row.headers.first], associate_header_values(row)]
      end.to_h
    end

    def associate_header_values(row)
      collection = headers(row).map{ |header| { header => row[header] } }
      collection.select!{|option| option.keys.any? }
    end

    def headers(row)
      row.headers[1, row.headers.length]
    end

  end

end

#convert a xlsx sheets into json files
# Convert.to_json("demo.xlsx")

#convert a xlsx sheets into csv files
# Convert.to_csv("demo.xlsx")
