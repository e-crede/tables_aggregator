require 'date'
require 'sqlite3'
require 'xsv'
require 'digest'

class ClusterField
    def initialize(config)
        # Replace TOKENS with actual file names
        name = config[:file_name]
        name.gsub!("$(YEAR)", Date.today.strftime("%Y"))
        name.gsub!("$(YEAR_DIGIT)", Date.today.strftime("%Y")[3])
        name.gsub!("$(WEEK)", Date.today.strftime("%V"))
        name.gsub!("$(MONTH)", Date.today.strftime("%m"))
        name.gsub!("$(DAY)", Date.today.strftime("%d"))
        config[:file_name] = name

        # Assign instance variables against every config element
        @file_name = config[:file_name]
        @start_row = config[:start_row]
        @end_row = config[:end_row]
        @sheet_name = config[:sheet_name]
        @column_names = config[:column_names]
        @store_name = config[:store_name]
        @db_table = config[:db_table]
        @save_to_db = config[:save_to_db]
        @archive_file = config[:archive_file] 

        # Local instance variables
        # TODO: import as yaml config
        @db_name = "db/test.db"
        @tmp_dir = "tmp/"
    end

    def grab
        path = @file_name
        if File::exists?(path)
            mtime = File::mtime(path)
            puts "INFO: Found match: #{path}, last modified on #{mtime}"
            File.copy_stream(path,@tmp_dir << @store_name )
        else
            abort("ERROR: File not found: #{path}")
        end
     end

     def archive
        err = false
        today = "db/#{Date.today.to_s}"
        new_file = "tmp/#{@store_name}"
        old_file = "#{today}/#{@store_name}"
        puts new_file
        puts old_file
        Dir.mkdir(today) unless Dir.exist? (today)
        # Don't save again if it's the same day and files are identical
        checksum_old = Digest::MD5.hexdigest(File.read(old_file))
        checksum_new = Digest::MD5.hexdigest(File.read(new_file))
        if checksum_new != checksum_old
            puts "Copying to archive directory: #{today}"
            Dir.mkdir(today) unless Dir.exist? (today)
            File.copy_stream(new_file, today << "/" << @store_name)
        else
            puts "WARNING: Identical file already backed up on the same day"
            err = true
        end
        return err
    end

    def excel_array
        puts "INFO: Reading excel file into memory..."
        file_name = "tmp/" << @store_name
        x = Xsv.open(file_name)
        #TODO - Include first within an array selection
        sheet = x.sheets_by_name(@sheet_name).first
        sheet.parse_headers!
        sheet
    end

    # TODO: Optimise and simplify method
    def save_to_db
        return puts "WARNING: Saving to DB disabled" unless @save_to_db
        arr = excel_array
        db = SQLite3::Database.open @db_name
        today = Date.today.to_s
        db.execute "CREATE TABLE IF NOT EXISTS #{@db_table}(date TEXT)"
        db.results_as_hash = true
        
        # Verify that columns in config exist in excel
        @column_names.each{ |col_config|
            found = false
            arr[0].each { |col_arr,val|
                found = true unless col_config[col_arr.to_sym] == nil
            }
            abort("ERROR: Column '#{col_config.keys[0]}' present in config, but not in the data array!") unless found
        }
    
        # Lambda to cleanup column names
        column_name_format = lambda { |i|
            old = i
            column_formatted = i.keys[0].to_s
            column_formatted.strip!
            column_formatted.force_encoding("ASCII")
            column_formatted.gsub!(" ","_")
            column_formatted
        }
    
        # Add new columns to the table if they don't exist yet
        @column_names.each{ |col_name|
            column_formatted = column_name_format.call(col_name)
            records = db.execute  "SELECT COUNT(*) AS CNTREC FROM pragma_table_info('#{@db_table}') WHERE name='#{column_formatted}'"
            db.execute "ALTER TABLE #{@db_table} ADD #{column_formatted} #{col_name.values[0]} DEFAULT ''" if records[0][0] == 0
        }
    
        # Prepare column names for SQL Query
        col_names_formatted = ""
        col_names_list = []
        @column_names.each{ |col_name|
            column_formatted = column_name_format.call(col_name)
            col_names_formatted = col_names_formatted + "'#{column_formatted}' "
            col_names_list.push(column_formatted)
        }
        col_names_formatted.strip!
        col_names_formatted.gsub!(" ",",")
        col_names_formatted << ',date'
        
        # Iterate through data rows and execute SQL query that inserts data into the db
        today = Date.today.to_s
        arr.each{|row|
            values_string = ""
            col_names_list.each{ |key|
            values_string = values_string + "'#{row[key]}' "
            }
            values_string.strip!
            values_string.gsub!(" ",",")
            values_string = "#{values_string},'#{Date.today.to_s}'" # append date to raw values
            db.execute "INSERT INTO #{@db_table} (#{col_names_formatted}) VALUES (#{values_string})"
        }    
    end

    def cleanup
    #TODO: delete tmp files
    end 
end


# Initialise program directory structure
def init_structure
    dirs = {
        tmp: "tmp",
        db: "db",
        db_src: "db/src",
    }
    dirs.each { |i,j| Dir.mkdir(j) unless Dir.exist? j}
end



# Files configuration
# file_name accepts tokens, e.g. $(YEAR) to dynamically update file paths
# See ClusterField class initialize method for details
configurations = [
    {
        file_name: "test/my_file.xlsx",
        start_row: 1,
        end_row: 10000,
        sheet_name: "Sheet1",
        column_names: [{"ColumnA": "TEXT"},{"ColumnC": "TEXT"}, {"ColumnD": "TEXT"}],
        store_name: "my_file_2022.xlsx",
        db_table: "my_table",
        save_to_db: true
    },
    {
        file_name: "test/my_file2.xlsx",
        start_row: 1,
        end_row: 10000,
        sheet_name: "Sheet1",
        column_names: [{"ColumnA": "TEXT"},{"ColumnC": "TEXT"}, {"ColumnD": "TEXT"}],
        store_name: "my_file_2022.xlsx",
        db_table: "my_table",
        save_to_db: false
    }
]

# Main program flow
init_structure
configurations.each{|config|
    puts "INFO: Begin processing: #{config}"
    data = ClusterField.new(config)
    puts "INFO: Retrieving file..."
    data.grab
    puts "INFO: Archiving file..."
    data.archive
    puts "INFO: Saving file to database..."
    data.save_to_db
    puts "INFO: Cleaning up...\n\n"
    data.cleanup
}

