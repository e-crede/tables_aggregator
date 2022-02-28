require 'date'
require 'sqlite3'
require 'xsv'

 # Load configuration
 def load_configuration
    config = [
        {
            file_name: "test/my_file.xlsx",
            start_row: 1,
            end_row: 10000,
            sheet_name: "Sheet1",
            column_names: [{"ColumnA": "TEXT"},{"ColumnC": "TEXT"}, {"ColumnD": "TEXT"}],
            store_name: "my_file_2022.xlsx",
            db_table: "my_table",
            save_to_db: true
        }
    ]
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

# TODO: Retrieve latest file backup time
def check_backup_time(store_name)
    bak_mtime=Time.mktime(2000)
end

 # Perform initial file checks and copy to a local directory
  def check_and_get_file(element,bak_mtime)
    path = element[:file_name]
    tmp_dir = "tmp/"
    if File::exists?(path)
        mtime = File::mtime(path)
        puts "Matching file found: #{path}, last modification on #{mtime}"
        File.copy_stream(path,tmp_dir << element[:store_name] ) if mtime > bak_mtime
    else
        puts "File not found: #{path}"
    end
 end

# Create archive dir & backup new file
 def archive_file(element)
    new_file = "tmp/" << element[:store_name]
    old_file = "test/old_file.txt"
    unless File.identical?(new_file,old_file)
        today = "db/" << Date.today.to_s
        puts "Copying to archive directory: #{today}"
        Dir.mkdir(today) unless Dir.exist? (today)
        File.copy_stream(new_file, today << "/" << element[:store_name])
    else
        puts "New backup identical to old, discard #{new_file}"
        File::delete(new_file)
    end
end

# Create an excel array
# each array element is a key value pair,
# where key is header and value is a row value
def excel_array(element)
    puts "Reading excel file into memory"
    file_name = "tmp/" << element[:store_name]
    x = Xsv.open(file_name)
    sheet = x.sheets_by_name(element[:sheet_name]).first
    sheet.parse_headers!
    sheet
end

 # Append database table
 def save_to_db(element)
    arr = excel_array(element)
    db_name = "db/test.db"
    db = SQLite3::Database.open db_name
    db_table = element[:db_table]
    today = Date.today.to_s
    
    db.execute "CREATE TABLE IF NOT EXISTS #{db_table}(sample_1 TEXT)"
    db.results_as_hash = true

    
    # Verify that columns in config exist in excel
    # FIXME: should abort the complete method
    element[:column_names].each{ |col_config|
        found = false
        arr[0].each { |col_arr,val|
            found = true unless col_config[col_arr.to_sym] == nil
        }
        puts "ERROR: '#{col_config.keys[0]}' present in config, but not in data array!" unless found
    }

    # cleanup column names
    column_name_format = lambda { |i|
        old = i
        column_formatted = i.keys[0].to_s
        column_formatted.strip!
        column_formatted.force_encoding("ASCII")
        column_formatted.gsub!(" ","_")
        column_formatted
    }

    # add new columns to the table if they don't exist yet
    element[:column_names].each{ |col_name|
        column_formatted = column_name_format.call(col_name)
        records = db.execute  "SELECT COUNT(*) AS CNTREC FROM pragma_table_info('#{db_table}') WHERE name='#{column_formatted}'"
        db.execute "ALTER TABLE #{db_table} ADD #{column_formatted} #{col_name.values[0]} DEFAULT ''" if records[0][0] == 0
    }

    # Build SQL INSERT INTO <tables> (COLUMNS) VALUES (Values) Statement
    # Column names
    col_names_formatted = ""
    col_names_list = []
    element[:column_names].each{ |col_name|
        column_formatted = column_name_format.call(col_name)
        col_names_formatted = col_names_formatted + "'#{column_formatted}' "
        col_names_list.push(column_formatted)
    }
    col_names_formatted.strip!
    col_names_formatted.gsub!(" ",",")
    #puts "Column Names SQL: #{col_names_formatted}"
    #puts "Column Names list: #{col_names_list}"
    
    # Iterate through rows and insert data into the database
    arr.each{|row|
        values_string = ""
        col_names_list.each{ |key|
        values_string = values_string + "'#{row[key]}' "
        }
        values_string.strip!
        values_string.gsub!(" ",",")
        db.execute "INSERT INTO #{db_table} (#{col_names_formatted}) VALUES (#{values_string})"
    }    
end
  
 # process files
 def process_files
    config = load_configuration
    count = 0
    config.each {|element|
        puts "[File_#{count}] Start"
        puts "Element: #{element}"
        bak_mtime = check_backup_time(element)
        check_and_get_file(element,bak_mtime)
        archive_file(element)
        save_to_db(element)
        count+=1
    }
 end

# Main program flow
init_structure
process_files
