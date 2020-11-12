require "byebug"
RSpec.describe RankmiIndicatorsUpdate do
  it "has a version number" do
    expect(RankmiIndicatorsUpdate::VERSION).not_to be nil
  end

#-------------------------1------------------------
  it "copy file for first step" do
    FileUtils.cp("excel_files/excel_copy.xlsx", "excel_files/excel.xlsx")
    expect(File.exist?("excel_files/excel.xlsx")).to eq(true)
  end

  it "check if file for first step is created" do
    RankmiIndicatorsUpdate::Parse.transform_to_rankmi_template('excel_files/excel.xlsx')
    expect(File.exist?("Metas BCP.xlsx")).to eq(true)
  end
  
  it "Match the contents for excel file of step 1" do 
    expected = RubyXL::Parser.parse('excel_files/Metas BCP.xlsx')
    output = RubyXL::Parser.parse('Metas BCP.xlsx')
    expected[0].each_with_index do |expected_val, i|
        rows= expected[0][0].size
        j=0
        while j<rows do
            expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
            j+=1
        end
    end
  end

#-------------------------2------------------------

  it "copy file for second step" do
    FileUtils.cp("excel_files/excel-2_copy.xlsx", "excel_files/excel-2.xlsx")
    expect(File.exist?("excel_files/excel-2.xlsx")).to eq(true)
  end

  it "check if file for second step is created" do
    RankmiIndicatorsUpdate::Parse.first_output_of_step_two('excel_files/excel-2.xlsx')
    expect(File.exist?("Metas BCP con info rankmi y tipo.xlsx")).to eq(true)
  end
  
  it "Match the contents for excel file of step 2 file 1" do 
    expected = RubyXL::Parser.parse('excel_files/Metas BCP con info rankmi y tipo.xlsx')
    output = RubyXL::Parser.parse('Metas BCP con info rankmi y tipo.xlsx')
    expected[0].each_with_index do |expected_val, i|
        rows= expected[0][0].size
        j=0
        while j<rows do
            expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
            j+=1
        end
    end
  end
#-------------------------3------------------------
it "check file for second  step file 2" do
  expect(File.exist?("excel_files/Metas BCP con info rankmi y tipo.xlsx")).to eq(true)
end

it "check if file for second step is created" do
  RankmiIndicatorsUpdate::Parse.separate_records('excel_files/Metas BCP con info rankmi y tipo.xlsx')
  expect(File.exist?("Metas a crear.xlsx")).to eq(true)
end

it "Match the contents for excel file of step 2 file 2" do 
  expected = RubyXL::Parser.parse('excel_files/Metas a crear.xlsx')
  output = RubyXL::Parser.parse('Metas a crear.xlsx')
  expected[0].each_with_index do |expected_val, i|
      rows= expected[0][0].size
      j=0
      while j<rows do
          expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
          j+=1
      end
  end
end

#-------------------------4------------------------
it "check file for second  step file 3,4" do
  expect(File.exist?("excel_files/Metas BCP con info rankmi y tipo.xlsx")).to eq(true)
end

it "check if file for second step is created" do
  RankmiIndicatorsUpdate::Parse.define_goals_to_eliminate_and_update('excel_files/Metas BCP con info rankmi y tipo.xlsx')
  expect(File.exist?("Análisis accione.xlsx")).to eq(true)
  expect(File.exist?("Metas a borrar.xlsx")).to eq(true)
end

it "Match the contents for excel file of step 2 file 4" do 
  expected = RubyXL::Parser.parse('excel_files/Análisis accione.xlsx')
  output = RubyXL::Parser.parse('Análisis accione.xlsx')
  expected[0].each_with_index do |expected_val, i|
      rows= expected[0][0].size
      j=0
      while j<rows do
          expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
          j+=1
      end
  end
end
it "Match the contents for excel file of step 2 file 3" do 
  expected = RubyXL::Parser.parse('excel_files/Metas a borrar.xlsx')
  output = RubyXL::Parser.parse('Metas a borrar.xlsx')
  expected[0].each_with_index do |expected_val, i|
      rows= expected[0][0].size
      j=0
      while j<rows do
          expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
          j+=1
      end
  end
end
#-------------------------5------------------------
it "check file for second  step file 5" do
  expect(File.exist?("excel_files/Metas BCP con info rankmi y tipo.xlsx")).to eq(true)
end

it "check if file for second step is created" do
  RankmiIndicatorsUpdate::Parse.excel_with_goals_to_update('excel_files/Metas BCP con info rankmi y tipo.xlsx')
  expect(File.exist?("Metas a actualizar.xlsx")).to eq(true)
end

it "Match the contents for excel file of step 2 file 4" do 
  expected = RubyXL::Parser.parse('excel_files/Metas a actualizar.xlsx')
  output = RubyXL::Parser.parse('Metas a actualizar.xlsx')
  expected[0].each_with_index do |expected_val, i|
      rows= expected[0][0].size
      j=0
      while j<rows do
          expect(expected_val[j]&.value).to eq(output[0][i][j]&.value)
          j+=1
      end
  end
end

#----------------del files
it "Clean up files" do 
  a = 'Análisis accione.xlsx'
  b = 'Metas a actualizar.xlsx'
  c = 'Metas a borrar.xlsx'
  d = 'Metas a crear.xlsx'
  e = 'Metas BCP.xlsx'
  f = 'Metas BCP con info rankmi y tipo.xlsx'
 
  File.delete(a) if File.exists? a
  expect(File.exist?(a)).to eq false
  File.delete(b) if File.exists? b
  expect(File.exist?(b)).to eq false
  File.delete(c) if File.exists? c
  expect(File.exist?(c)).to eq false
  File.delete(d) if File.exists? d
  expect(File.exist?(d)).to eq false
  File.delete(e) if File.exists? e
  expect(File.exist?(e)).to eq false
  File.delete(f) if File.exists? f
  expect(File.exist?(f)).to eq false
end
end
