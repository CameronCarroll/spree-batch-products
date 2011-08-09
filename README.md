BatchProducts
=============

An extension aimed at providing the ability to create or update collections of products or variants.

BatchProducts depends on the [Spreadsheet](http://rubygems.org/gems/spreadsheet "Spreadsheet") gem to process uploaded excel files which are stored using Paperclip.  If DelayedJob is detected, the process of uploading a datasheet enqueues the datasheet for later processing.  If not, the datasheet is processed when it is uploaded.

Each ProductDatasheet record has 4 integer fields that give a basic description of the datasheet's effectiveness:

* `:matched_records =>` sum of records that were matched by each row
* `:updated_records =>` sum of all records updated by each row
* `:failed_records =>` sum of all records that had a 'false' return when saved
* `:failed_queries =>` number of rows that matched no Product or Variant records

Installation
============

To incorporate the BatchProducts extension into your Spree application, add the following to your gemfile depending on your version of Spree:

For Spree versions 0.50 and above, use:
`gem 'spree_batch_products', :git => 'git://github.com/minustehbare/spree-batch-products.git', :branch => '0-50-stable'`

For Spree versions 0.60 and above, use:
`gem 'spree_batch_products', :git => 'git://github.com/minustehbare/spree-batch-products.git', :branch => '0-60-stable'`

Follow it up with a `bundle install`.

When your bundle has finished, mirror the assets and migrations into your migrations folder with `rake spree_batch_products:install` and then run `rake db:migrate`.  This will create the ProductDatasheet(s) model and database table along with the handy statistic fields listed above.

If you are using DelayedJob, the Jobs table should already be created, or it will be created when you install DelayedJob.

Having done these things, you can log into the admin interface of your application and click on the 'Products' tab.  Listed as a sub-tab you'll see 'Batch Updates'.  This is where you can upload a new spreadsheet for processing or view existing spreadsheets that have already been completed or are pending to be processed.

Example
=======

ProductDatasheets rely on a few assumptions: the first row defines the attributes of the records you want to update, and the first (and possibly second) cell of that row defines the attribute to search records by.   

Consider a simple datasheet:

![](/minustehbare/spree-batch-products/raw/master/example/sample_spreadsheet.png)

Notice that the first cell defines the search attribute as `:sku`.  Since this attribute is exclusive to the Variant model, it is a 'collection' of variants that we are updating.  The second attribute that is defined is `:price`.  

Ideally, the first row of the datasheet will contain _all_ of the attributes that belong to the model you're updating but it is only necessary to reference the ones that you will be updating.  In this case, we are only updating the `:price` attribute on any records we find.

The second row and on define the 'queries' that are executed to retrieve a collection of records.  The first (and only) row translates to `Variant.where(:sku => 'ROR-00014').all`.  Each record in the collection executes an `#update_attributes(attr_hash)` call where `attr_hash` is defined by the remaining contents of the row.  Here the attributes hash is `{:price => 902.10}`.

If a query returns no records, or if the search attribute does not belong to Variant or Product attributes then it is reported as 'failed'.  Any records matched by the query are added to the `:matched_records` attribute of the datasheet.  Records that have a `true` return on the `#update_attributes(attr_hash)` call are added to the `:updated_records` attribute and those that have a `false` return are added to the `:failed_records` attribute.

Record Creation: Products
-------------------------

To create Product records through a ProductDatasheet the first row must define `:id` as the search attribute.  Each row should have an empty value for the `:id` column otherwise Product records will be located by the value supplied.  Record creation succeeds so long as the `:name`, `:permalink`, and `:price` attributes on each row are defined.

Product records must be defined on a separate sheet from your variants. I don't have time to make it perfect, allowing both definitions on the same page.

Record Creation: Variants
-------------------------

In order to define Variant record creation, you need to make a second spreadsheet in your excel workbook: The first sheet is to list products, whereas the second sheet lists variants.

To create Variant records, the extension looks for a product_id column AFTER the blank 'id' column. For any record creation, either product or variant, you MUST have a blank 'id' column in the first cell of your row. The product_id field is used to associate a variant with its parent product. If you already know the ID of the product for a given variant, you can define it explicitly (in integer form.) If you don't know the ID already, you can also provide the product's name to search by.

Note that variants are dependent upon upon type definitions within a product: Record creation will not succeed unless you, at the very least, define an option type in the Option_Types column of your sheet. An example of this would simply be Color:

Record Creation: Option Types & Option Values
---------------------------------------------

Options are added to their respesctive record type, but are defined together on the Variant sheet. After creating the product and its variant, the program proceeds to handle exceptions, including option types. Option types themselves are pulled out of a product's exception_hash and associated to the variant's parent product.

There are a couple syntactic elements to keep in mind for definining option types: The regular expressions used are "option_type_regex = /\w*:/" & "option_value_regex = /(\w*,)|(\w*;)/" respectively. I suggest your test your data out in something like [rubular](http://rubular.com/) or any given alternative to be sure that what you have will be matched. (Or if you're a regex wizard, you could make the query more robust...)
The above expressions are designed to parse a string as such: "Color:blue; Size:small;" --- Option_type_regex will yield "Color: and Size:" Option_value_regex will yield each option (red, blue, green; small, large;) in an array. Your option values will not be picked up unless they are terminated with either a comma or semicolon. Your option types will not be picked up unless they are terminated with a colon. Finally, separate option type/value trees with a space.

Due to not really understanding variants and options, I prepared the system to use multiple option values for a single type, for a single variant. I realize that this doesn't render prettily by default, but if you so wish, you can define something like Color:blue,green,red and it'll render in the variants partial as Color: blue, Color: green, Color: red

Record Updating
---------------

Updating collections of records follows similarly from the example.  Updating Product collections requires a search attribute that is present as an attribute column on the Products table in the database; the same is true for Variant collections.  Attributes with empty value cells are not included in the attributes hash to update the record.

Copyright (c) 2011 minustehbare, released under the New BSD License

Contributions by Sanarothe, referencing [spree-import-products](https://github.com/joshmcarthur/spree-import-products/) and [ar-loader](https://github.com/autotelik/AR-Loader) -- If I couldn't figure out how to do something, I probably looked at how you did it. <3
