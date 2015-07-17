require 'simple_xlsx_reader'
require 'csv'
require 'json'
require 'pry'

module Xlsx
  class Convert
    attr_accessor :file, :output_dir

    def initialize
      @file = nil
      @output_dir = '/tmp/'
    end

    def to_json
      self.class.to_json(@file, self)
    end

    def to_csv
      self.class.to_json(@file, self)
    end

    def build_file_name(string, format=".csv")
      "#{@output_dir}#{string.gsub(/\s+/, '_').downcase}#{format}"
    end

    class << self

      def to_json(file, obj = self.new)
        xlsx = file_content_to_csv(file, obj)
        csv_files_to_json(xlsx.sheets.map(&:name), obj)
      end

      def to_csv(file, obj = self.new)
        file_content_to_csv(file, obj)
      end

      private
      def file_content_to_csv(file, obj = self.new)

        xlsx = SimpleXlsxReader.open(file)

        xlsx.sheets.each do |sheet|
          file_name = obj.build_file_name(sheet.name)
          File.open(file_name,  'w') do |f|
            sheet.rows.each do |row|
              f.write "#{row.join(',')}\n"
            end
          end
        end

        xlsx

      end

      def csv_files_to_json(sheets_names, obj)
        sheets_names.each do |sheet_name|
          associated_data = build_hash_object(sheet_name, obj)
          json_content    = build_json_data(associated_data)

          File.open(obj.build_file_name(sheet_name, '.json'), 'w') do |csv|
            csv.write json_content
          end

          File.delete(obj.build_file_name(sheet_name))
        end
      end

      def build_json_data(data)
        JSON.pretty_generate(data)
      end

      def build_hash_object(sheet_name, obj)
        CSV.foreach(obj.build_file_name(sheet_name), headers: true).map do |row|
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
end
