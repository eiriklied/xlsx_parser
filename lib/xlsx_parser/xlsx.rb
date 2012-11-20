require 'tmpdir'
require 'zip/zipfilesystem'

class Xlsx
  def initialize(filename)
    @tmpdir = Dir.mktmpdir
    extract_zip_content(filename)
  end



private
  def extract_zip_content(zipfilename)
    Zip::ZipFile.open(zipfilename) do |zip|
      process_zipfile(zipfilename,zip)
    end
  end

  def process_zipfile(zipfilename, zip, path='')
    @sheet_files = []
    Zip::ZipFile.open(zipfilename) {|zf|
      zf.entries.each {|entry|
        #entry.extract
        if entry.to_s.end_with?('workbook.xml')
          open(@tmpdir+'/workbook.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.end_with?('sharedStrings.xml')
          open(@tmpdir+'/sharedStrings.xml','wb') {|f|
            f << zip.read(entry)
          }
        end
        if entry.to_s.end_with?('styles.xml')
          open(@tmpdir+'/styles.xml','wb') {|f|
            f << zip.read(entry)
          }
        end 
        if entry.to_s =~ /sheet([0-9]+).xml$/
          nr = $1
          open(@tmpdir+'/'+@file_nr.to_s+"sheet#{nr}.xml",'wb') {|f|
            f << zip.read(entry)
          }
          @sheet_files[nr.to_i-1] = @tmpdir+'/'+@file_nr.to_s+"sheet#{nr}.xml"
        end
      }
    }
    return
  end

end