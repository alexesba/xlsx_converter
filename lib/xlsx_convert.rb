require 'simple_xlsx_reader'
require 'csv'
require 'json'

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
      self.class.to_csv(@file, self)
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
              f.write "#{row.map{|field| field.to_s.gsub('"', '""').gsub(/\n/, ' ').gsub('  ', ' ')}.join("||")}\n"
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
        CSV.foreach(obj.build_file_name(sheet_name), headers: true, col_sep: "||", quote_char:"\0", force_quotes: true).map do |row|
          header = normalize_header(row.headers.first)
          header_name = normalize_header(row[header])
          [ header_name, associate_header_values(row)] unless header_name.empty?
        end.compact.to_h
      end

      def normalize_header(header)
        header.to_s.gsub('.0', '').strip
      end

      def associate_header_values(row)
        collection = headers(row).map{ |header| { normalize_header(header) => row[header] } }
        collection = collection.select {|option| option.keys.any? && option.values.any? }
        collection.reduce({}) {|h,pairs| pairs.each {|k,v| h[normalize_header(k)] =  v}; h}
      end

      def headers(row)
        row.headers[1, row.headers.length]
      end

    end

  end
end
