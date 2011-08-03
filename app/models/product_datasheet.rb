class ProductDatasheet < ActiveRecord::Base
  require 'spreadsheet'

has_attached_file :xls, :path => ":rails_root/uploads/product_datasheets/:id/:basename.:extension"

validates_attachment_presence :xls
validates_attachment_content_type :xls, :content_type => ['application/vnd.ms-excel','text/plain']

scope :not_deleted, where("product_datasheets.deleted_at is NULL")
scope :deleted, where("product_datasheets.deleted_at is NOT NULL")

def path
  "#{Rails.root}/uploads/product_datasheets/#{self.id}/#{self.xls_file_name}"
end

#// Main method for control of spreadsheet processing:
#// Opens up a workbook and its two worksheets. The first sheet is used to define products,
#// and the second sheet is used to define variants.


def process
  uploaded_workbook = Spreadsheet.open self.path

product_sheet = uploaded_workbook.worksheet(0)
product_columns = [product_sheet.dimensions[2], product_sheet.dimensions[3]]
product_headers = product_sheet.row(0)
perform(product_sheet)

if not uploaded_workbook.worksheet(1).nil?
  variant_sheet = uploaded_workbook.worksheet(1)
  variant_columns = [variant_sheet.dimensions[2], variant_sheet.dimensions[3]]
  variant_headers = variant_sheet.row(0)
  perform(variant_sheet)
end

@records_matched = 0
@records_updated = 0
@records_failed = 0
@failed_queries = 0
end #process
  
  def perform(sheet)
    #// Passing 1 into each(1) defines how many rows to skip before processing the spreadsheet.
    #// Since the first row is already defined in headers=worksheet.row(0), we start at the data.
    sheet.each(1) do |row|

    #// A note on dimensions: values 0-3 of the return array contain the first used column, and the first
    #// unused column. See documentation at http://spreadsheet.rubyforge.org/Spreadsheet/Worksheet.html
    #// Load columns and headers for the new sheet:
      columns = [sheet.dimensions[2], sheet.dimensions[3]]
      headers = sheet.row(0)
      attr_hash = load_headers(row, columns, headers)

      #// Checks first header value for a blank 'id', which signifies record creation.
      #// If record is to be created, checks for product_id column, which signifies variant creation.
      
      if headers[0] == 'id' && row[0].nil? && headers[1] == 'product_id' && !row[1].nil?
        create_variant(attr_hash)
      elsif headers[0] == 'id' && row[0].nil?
        create_product(attr_hash)
      elsif Product.column_names.include?(headers[0])
        process_products(headers[0], row[0], attr_hash)
      elsif Variant.column_names.include?(headers[0])
        process_variants(headers[0], row[0], attr_hash)
      else
        #TODO do something when the batch update for the row in question is invalid
        @failed_queries = @failed_queries + 1
      end
      columns = nil
      headers = nil
    end #sheet.each

    attr_hash = { :processed_at => Time.now, 
                  :matched_records => @records_matched, 
                  :failed_records => @records_failed, 
                  :updated_records => @records_updated, 
                  :failed_queries => @failed_queries }
    self.update_attributes(attr_hash)
  end # perform
  
   #// Uses pre-defined headers array and associates each header value with its target value.
  #// Iterates between the first used and the first unused columns and grabs the header and row value.  
  def load_headers(row, columns, headers)
    attr_hash = {}
    for i in columns[0]..columns[1]
     attr_hash[headers[i]] = row[i] unless row[i].nil?
    end
    return attr_hash
  end
  
  #// Simply instantiates a new product using the attribute hash formed in load_headers
  def create_product(attr_hash)
    new_product = Product.new(attr_hash)
    @failed_queries = @failed_queries + 1 if not new_product.save
  end
  
  
  #// create_variant uses product_id in attr_hash:
  #// accepts string, integer values (string for lookup, integer for direct association.)
  #// If product is found, injects its ID into attr_hash in place of name
  def create_variant(attr_hash)
    product_to_reference = Product.where(headers[1] => attr_hash[headers[1]])
    if not product_to_reference.nil?
      attr_hash[headers[1]] = product_to_reference.id
    else
      #Fail miserably :/
    end
    new_variant = Variant.new(attr_hash)
    @failed_queries = @failed_queries + 1 if not new_product.save
  end
  
  def process_products(key, value, attr_hash)
    products_to_update = Product.where(key => value).all
    @records_matched = @records_matched + products_to_update.size
    products_to_update.each { |product| 
        if product.update_attributes attr_hash
            @records_updated = @records_updated + 1
        else
            @records_failed = @records_failed + 1
         end }
    @failed_queries = @failed_queries + 1 if products_to_update.size == 0
  end
  
  def process_variants(key, value, attr_hash)
    variants_to_update = Variant.where(key => value).all
    @records_matched = @records_matched + variants_to_update.size
    variants_to_update.each { |variant| 
        if variant.update_attributes attr_hash
           @records_updated = @records_updated + 1
        else
           @records_failed = @records_failed + 1
        end }
    @failed_queries = @failed_queries + 1 if variants_to_update.size == 0
  end
  
  def processed?
    !self.processed_at.nil?
  end
  
  def deleted?
    !self.deleted_at.nil?
  end
end
