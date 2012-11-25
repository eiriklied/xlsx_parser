require 'tmpdir'
require 'zip/zipfilesystem'
require 'nokogiri'

class Xlsx
  def initialize(filename)
    @sheets = []
    @last_column = 1 # default value to grow for each sheet
    @tmpdir = Dir.mktmpdir
    extract_zip_content(filename)
    #parse_shared_strings
    parse_workbook
    #parse_sheets
  end

  # An array of sheet names
  def sheets
    @sheet_names ||= @workbook_doc.
      xpath("//a:sheet", {"a" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}).
      collect do |sheet|
        sheet.attributes['name'].value
      end
  end


private
  
  def parse_workbook
    @workbook_doc = Nokogiri::XML(File.read("#{@tmpdir}/workbook.xml"))
  end

  def parse_shared_strings
    doc = Nokogiri::XML(File.read("#{@tmpdir}/sharedStrings.xml"))
    elements = doc.xpath('//a:t', {"a" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main"})
    
    @shared_strings = elements.collect{|node| node.text}
    # release for gc
    elements = nil 
    #but return something excpected :)
    @shared_strings
  end

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