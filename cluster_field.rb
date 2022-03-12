# ClusterField:
# A Script to grab excel files from remote (network) shares,
# copy them to local directories and copy data into an SQLite database

# frozen_string_literal: false

require 'date'
require 'sqlite3'
require 'xsv'
require 'digest'

# Handle file I/O operations
class ClusterFiles
  def initialize(config)
    # Assign instance variables against every config element
    @file_name = replace_tokens(config[:file_name])
    @start_row = config[:start_row]
    @sheet_name = config[:sheet_name]
    @column_names = config[:column_names]
    @store_name = config[:store_name]
    @db_table = config[:db_table]
    @save_to_db = config[:save_to_db]
    @archive_file = config[:archive_file]
    # TODO: import as yaml config
    @db_name = 'db/test.db'
    @tmp_dir = 'tmp/'
    @today_dir = "db/#{Date.today}"
    @new_file = "tmp/#{@store_name}"
    @old_file = "#{@today_dir}/#{@store_name}"
  end

  def grab
    err = false
    if File.exist?(@file_name)
      mtime = File.mtime(@file_name)
      puts "INFO: Found match: #{@file_name}, last modified on #{mtime}"
      File.copy_stream(@file_name, @tmp_dir << @store_name)
    else
      puts "ERROR: File not found: #{@file_name}"
      err = true
    end
    err
  end

  def archive
    err = false
    return err unless @archive_file

    Dir.mkdir(@today_dir) unless Dir.exist? @today_dir
    if checksums_differ(@old_file, @new_file)
      File.copy_stream(@new_file, @today_dir << '/' << @store_name)
    else
      puts "WARNING: Identical file with backup, won't copy"
      err = true
    end
    err
  end

  def cleanup
    # TODO: delete tmp files
    puts "INFO: Cleaning up...\n\n"
  end

  private

  # Replace TOKENS with actual file names
  def replace_tokens(name)
    today = Date.today
    name.gsub!('$(YEAR)', today.strftime('%Y'))
    name.gsub!('$(YEAR_DIGIT)', today.strftime('%Y')[3])
    name.gsub!('$(WEEK)', today.strftime('%V'))
    name.gsub!('$(MONTH)', today.strftime('%m'))
    name.gsub!('$(DAY)', today.strftime('%d'))
    name
  end

  def checksums_differ(old_file, new_file)
    checksum_old = Digest::MD5.hexdigest(File.read(old_file)) if File.exist?(old_file)
    checksum_new = Digest::MD5.hexdigest(File.read(new_file))
    checksum_new != checksum_old
  end
end

# Handles Excel Tables transfer into the Database
class ClusterField < ClusterFiles
  def initialize(config)
    super
    @db = SQLite3::Database.open @db_name
  end

  private

  def create_table
    @db.execute "CREATE TABLE IF NOT EXISTS #{@db_table}(date TEXT, datetime TEXT)"
    @db.results_as_hash = true
  end

  # Verify that columns in config exist in excel
  def check_columns(arr)
    @column_names.each do |col_config|
      found = false
      arr[0].each do |col_arr, _val|
        found = true unless col_config[col_arr.to_sym].nil?
      end
      abort("ERROR: Column '#{col_config.keys[0]}' present in config, but not in the data array!") unless found
    end
  end

  def normalise_name(col_name)
    col_name.strip!
    col_name.force_encoding('ASCII')
    col_name.gsub!(' ', '_')
    col_name
  end

  # Add new columns to the table if they don't exist yet
  def add_new_columns
    @column_names.each do |col_name|
      col_norm = normalise_name(col_name.keys[0].to_s)
      records = @db.execute "SELECT COUNT(*) AS CNTREC FROM pragma_table_info('#{@db_table}') WHERE name='#{col_norm}'"
      @db.execute "ALTER TABLE #{@db_table} ADD #{col_norm} #{col_name.values[0]} DEFAULT ''" if records[0][0].zero?
    end
  end

  def prepare_columns_for_sql
    col_names_formatted = ''
    col_names_list = []
    @column_names.each do |col_name|
      column_formatted = normalise_name(col_name.keys[0].to_s)
      col_names_formatted += "'#{column_formatted}' "
      col_names_list.push(column_formatted)
    end
    col_names_formatted.strip!.gsub!(' ', ',') << ',date,datetime'
    [col_names_formatted, col_names_list]
  end

  def insert_rows(arr, col_names_formatted, col_names_list)
    arr.each do |row|
      values_string = ''
      col_names_list.each do |key|
        values_string += "'#{row[key]}' "
      end
      values_string.strip!
      values_string.gsub!(' ', ',')
      values_string = "#{values_string},'#{Date.today}','#{DateTime.now}'" # append date to raw values
      @db.execute "INSERT INTO #{@db_table} (#{col_names_formatted}) VALUES (#{values_string})"
    end
  end

  def excel_array
    puts 'INFO: Reading excel file into memory...'
    file_name = 'tmp/' << @store_name
    x = Xsv.open(file_name)
    sheet = x.sheets_by_name(@sheet_name).first
    sheet.row_skip = @start_row
    sheet.parse_headers!
    sheet
  end

  def check_for_duplicates(arr, col_names_formatted, col_names_list)
    # TODO: check that fields are not in db already
  end

  public

  def save_to_db
    puts 'INFO: Saving file to database...'
    return puts 'WARNING: Saving to DB disabled' unless @save_to_db

    create_table
    add_new_columns
    arr = excel_array
    check_columns(arr)
    col_names_formatted, col_names_list = prepare_columns_for_sql
    check_for_duplicates(arr, col_names_formatted, col_names_list)
    insert_rows(arr, col_names_formatted, col_names_list)
  end
end

# Initialise program directory structure
def init_structure
  dirs = {
    tmp: 'tmp',
    db: 'db',
    db_src: 'db/src'
  }
  dirs.each { |_i, j| Dir.mkdir(j) unless Dir.exist? j }
end

# Files configuration
# file_name accepts tokens, e.g. $(YEAR) to dynamically update file paths
# See ClusterFile class initialize method for details
configurations = [
  {
    file_name: 'test/my_file.xlsx',
    start_row: 0,
    sheet_name: 'Sheet1',
    column_names: [{ 'ColumnA': 'TEXT' }, { 'ColumnC': 'TEXT' }, { 'ColumnD': 'TEXT' }],
    store_name: 'my_file_2022.xlsx',
    db_table: 'my_table',
    save_to_db: true,
    archive_file: true
  },
  {
    file_name: 'test/my_file.xlsx',
    start_row: 1,
    sheet_name: 'Sheet1',
    column_names: [{ 'textA': 'TEXT' }, { 'textB': 'TEXT' }],
    store_name: 'my_file2.xlsx',
    db_table: 'my_table_2',
    save_to_db: true,
    archive_file: false
  }
]

# Main program flow
init_structure
configurations.each do |config|
  puts "INFO: Begin processing: #{config}"
  data = ClusterField.new(config)
  err = data.grab
  err = data.archive unless err
  data.save_to_db unless err
  data.cleanup
end
