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
  @records_matched = 0
  @records_updated = 0
  @records_failed = 0
  @failed_queries = 0
  @skipped_records = 0

  product_sheet = uploaded_workbook.worksheet(0)
  perform(product_sheet)

  if not uploaded_workbook.worksheet(1).nil?
    variant_sheet = uploaded_workbook.worksheet(1)
    perform(variant_sheet)
  end
end #process
  
  def perform(sheet)
    #// Passing 1 into each(1) defines how many rows to skip before processing the spreadsheet.
    #// Since the first row is already defined in headers=worksheet.row(0), we start at the data.
    sheet.each(1) do |row|

    #// A note on dimensions: values 0-3 of the return array contain the first used column, and the first
    #// unused column. See documentation at http://spreadsheet.rubyforge.org/Spreadsheet/Worksheet.html
    #// Load columns and headers for the new sheet:
      columns = [sheet.dimensions[2], sheet.dimensions[3]]
      #// Captures the first row of the sheet, which contains the headers (model attributes)
      #// These headers are mostly used in the product/variant creation/update call, but there are some
      #// exceptions: option types, taxonomies and (possibly) 'ad-hoc option types'
      #// These exceptions are filtered in load_data and added to the exception list.
      raw_headers = sheet.row(0)
      #// header_return_array: 0 contains attr_hash, 1 contains exception_hash, 2 contains the filtered headers.
      header_return_array = load_data(row, columns, raw_headers)
      attr_hash = header_return_array[0]
      exception_hash = header_return_array[1]
      headers = header_return_array[2]
      
      #// Checks first header value for a blank 'id', which signifies record creation.
      #// If record is to be created, checks for product_id column, which signifies variant creation.
      #// note that create_variant requires the headers, but create_product does not.
      
      #// Checks for blank ID, non-blank product_id and a blank option_types.
      #// This is for creating variants if the product already has option types associated with it.
      if headers[0] == 'id' && row[0].nil? && headers[1] == 'product_id' && !row[1].nil?
        create_variant(attr_hash, headers, exception_hash)     
      #// Create products as normal
      elsif headers[0] == 'id' && row[0].nil? && headers[1] != 'product_id'
        create_product(attr_hash)
      elsif Product.column_names.include?(headers[0])
        process_products(headers[0], row[0], attr_hash)
      elsif Variant.column_names.include?(headers[0])
        process_variants(headers[0], row[0], attr_hash) 
      else
        #TODO do something when the batch update for the row in question is invalid
        @failed_queries += 1
      end
      columns = nil
      headers = nil
    end #sheet.each

    attr_hash = { :processed_at => Time.now, 
                  :matched_records => @records_matched, 
                  :failed_records => @records_failed, 
                  :updated_records => @records_updated, 
                  :failed_queries => @failed_queries}
    self.update_attributes(attr_hash)
  end # perform
  
  #// Exclusion hash is used to define columns that need to be processed separately. In order to define
  #// an exception, you need to hook it here to prevent it from being added to attr_hash, and also add
  #// a handler in the handle_exception method.
  def load_data(row, columns, headers)
  #// HEADER EXCLUSION LIST:
  #// ----------------------
      exclusion_list = [
        'Option_Types'
      ]
    attr_hash = {}
    exception_hash = {}
    sanitized_headers_array = []
    header_return_array = []
    
    for i in columns[0]..columns[1]
      exclusion_list.each do |exclusion|
        if headers[i] =~ /#{exclusion}/i
          exception_hash[exclusion] = row[i]
        elsif headers[i] == exclusion
          exception_hash[exclusion] = row[i]
        else
          attr_hash[headers[i]] = row[i] unless row[i].nil?
          sanitized_headers_array << headers[i]
        end
      end
     
    end
    header_return_array[0] = attr_hash
    header_return_array[1] = exception_hash
    header_return_array[2] = sanitized_headers_array
    return header_return_array
  end #load_data
  
  #// Add exception handlers to the case statement.
  def handle_exceptions(exception_hash, attr_hash)
    exception_hash.each do |exception_key, exception_value|
      
      if exception_value.nil?
        break
      end

      case exception_key
        #// Handle option types, which are only defined within variants. Uses the exception data to add option_types to the parent product 
        #// and then adds option_values to the variant.
        #// exception_package[0] is the exception_hash, which should contain currently only 'Option_Types'
        #// exception_package[1] is the attr_hash, which has the variant row data. 
      when 'Option_Types'
        individual_trees_regex = /\s/

        #// Variant should have had the ID injected rather than name already.
        parent_to_query = attr_hash['product_id']
        parent_product = Product.find_or_create_by_id(parent_to_query)
        
        #// if the variant exists already, find it by sku
        variant_to_query = attr_hash['sku']
        our_variant = Variant.find_by_sku(variant_to_query)
        
        #// Breaks the exception_value into individual option type/value trees for simplification of processing.
        #// Creates one option_type and a number of option_values for the product and variant respectively.
        option_trees = exception_value.split(individual_trees_regex)
        option_trees.each do |tree|
          #// options are returned as an array containing items, which are themselves arrays.
          #// option_types are arrays with only one value
          #// option_values have one of two possible values filled: The first handles commas, the latter, semicolons.
          option_return_array = parse_options(tree)
          option_type = option_return_array[0]
          option_values = option_return_array[1]
          #// Initialize parent product's option type.
          #// Yeah, I'm getting lazy. That parent_option shouldnt be global, but it is.
          parent_product.option_types = option_type.map do |type|
            type.gsub(':', '')
            OptionType.find_or_create_by_name_and_presentation(type, type.capitalize)
          end
          #// If the variant doesn't already exist, create it now that the parent product has option types.
          if our_variant.nil?
            our_variant = Variant.new(attr_hash)
            @failed_queries += 1 if not our_variant.save
          end 
          #// Get the parent option_type in scope:
          #// option_type array contains items, as arrays. It sucks, but that's what we get back from scan.
          parent_option = OptionType.find_by_name(option_type[0])
          #// Finally, associate option values with the variant.
          our_variant.option_values = option_values.map do |value|
            if !value.nil?
              OptionValue.find_or_create_by_name_and_presentation_and_option_type_id(value, value.capitalize, parent_option.id)
            else
              #Option values are nil. This shouldn't happen.
              @failed_queries += 1 
            end
          end  
        end #option_trees
      
      else
        #Exception not found
        @failed_queries += 1
      end
    end
  end #handle_exceptions
  
  #// Accepts a string of option_types and option_values in the form: "Color:red,blue,green;"
  #// Reduces this string to the parent option_type ("Color:") and child option_values ("red,blue,green;")
  def parse_options(option_string)
    option_type_regex = /\w*:/
    option_value_regex = /\w*;+/
    option_return_array = []
    #// Find the option_type and strip the colon out.
    #// Notice that option_types and option_values are stored in an array after the scan. For option_type,
    #// because there should only be one value per tree, always access value 0.
    #// Option values will either be in location 0 (First match) or location 1 (scond match), which gets commas and semicolons respectively.
    
    raw_option_type = option_string.scan(option_type_regex)
    option_type = raw_option_type[0].gsub(':', '')
    
    raw_option_values = option_string.scan(option_value_regex)
    raw_option_values.each do |value|
       value = value.gsub(';', '') 
    end
    #// Load return array
    option_return_array << option_type
    option_return_array << option_values
    return option_return_array
  end
  
  #// Simply instantiates a new product using the attribute hash formed in load_headers
  def create_product(attr_hash)
    product_already_exists = Product.find_by_sku(attr_hash['sku'])
    if product_already_exists
      @skipped_records += 1
    else
     new_product = Product.new(attr_hash)
     @failed_queries += 1 if not new_product.save
    end
  end
  
  
  #// create_variant uses product_id in attr_hash:
  #// accepts string, integer values (string for lookup, integer for direct association.)
  #// If product is found, injects its ID into attr_hash in place of name
  #// Note: handle_exceptions: 'option_types' creates the variant by itself.
  def create_variant(attr_hash, headers, exception_hash)
    product_to_reference = Product.find_by_name(attr_hash[headers[1]])
    if product_to_reference.nil?
      product_to_reference = Product.find_by_id(attr_hash[headers[1]])
    end
    if not product_to_reference.nil?
      attr_hash[headers[1]] = product_to_reference[:id]
    else
      @failed_queries = @failed_queries +1
      return
    end
    handle_exceptions(exception_hash, attr_hash)
    new_variant = Variant.find_by_sku(attr_hash['sku'])
    if new_variant.nil?
      new_variant = Variant.new(attr_hash)
      @failed_queries += 1 if not new_variant.save
    end
  end #create_variant
  
  #//
  def process_products(key, value, attr_hash)
    products_to_update = Product.where(key => value).all
    @records_matched = @records_matched + products_to_update.size
    products_to_update.each { |product| 
        if product.update_attributes attr_hash
            @records_updated = @records_updated + 1
        else
            @records_failed = @records_failed + 1
         end }
    @failed_queries += 1 if products_to_update.size == 0
  end #process_products
  
  #//
  def process_variants(key, value, attr_hash)
    variants_to_update = Variant.where(key => value).all
    @records_matched = @records_matched + variants_to_update.size
    variants_to_update.each { |variant| 
        if variant.update_attributes attr_hash
           @records_updated = @records_updated + 1
        else
           @records_failed = @records_failed + 1
        end }
    @failed_queries += 1 if variants_to_update.size == 0
  end #process_variants
  
  def processed?
    !self.processed_at.nil?
  end
  
  def deleted?
    !self.deleted_at.nil?
  end
end
