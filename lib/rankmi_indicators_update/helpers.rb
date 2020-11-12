module RankmiIndicatorsUpdate
  class Helper
    
    # add a cell if doesnt exist else add with a comma
    private_class_method def self.print(file,row,col,val)
      if file[row][col]&.value.to_s.empty?
          file&.add_cell(row,col,val) 
        else
          file[row][col]&.change_contents("#{file[row][col]&.value}, #{val}") 
        end
      rescue
        file&.add_cell(row,col,val) 
    end

    # find a column in a given file with a given column header name
    private_class_method def self.find_col(file, name)
      i=0
      found=nil
      while i<file[0]&.size.to_i
          cell = file[0][i]
          if cell&.value&.downcase==name&.downcase
              found=i
              break
          end
          i+=1
      end
      return found
    end

    # find a row in a shheet in a column with a header and the value to find in with current column prevention 
    private_class_method def self.find_row(file,column_name,value,curr_col)
      i=0
      found = nil
      column_index = find_col(file, column_name)
      return nil if !column_index
      while i < file.sheet_data.rows.size 
          break if !column_index
          if file[i][column_index]&.value&.downcase == value&.downcase && curr_col!=i
              found=i
              break
          end
          i+=1
      end
      return found
    end

# to copy an entire column from a sheet with a column name to a sheet in a column name and if all values need to append a value passs
  private_class_method def self.copy_column(from,from_column_name,to,to_column_name,append_str="")
    from_index = find_col(from, from_column_name )
    to_index = find_col(to, to_column_name )
    raise Exception.new("Either input file is missing '#{from_column_name}' or output file hasnt added the header '#{to_column_name}''") if !to_index || !from_index
    rows = from&.sheet_data&.rows&.size
    i=1
    while i<rows.to_i
      cell = from[i][from_index]
      print(to,i,to_index,append_str.empty? ? cell&.value : cell&.value.to_s+append_str)
      i+=1
    end
  end


  # add a header to a file pass an array of string in header param
  private_class_method def self.add_header_to_output(file,header)
    file&.sheet_name = "es"
    header.each_with_index do |val, i|
      print(file,0,i,val)

    end
  end
  
  
  # copy one value in entire column 
  private_class_method def self.copy_value_in_col(file,col,val)
    to_index = find_col(file, col )
    raise Exception.new("output file hasnt added the header '#{to_column_name}''") if !to_index 
    rows = file&.sheet_data&.rows&.size
    i=1
    while i<rows.to_i
      print(file,i,to_index,val)
      i+=1
    end
  end

  # create and insert col_description with the rules mentioned
  private_class_method def self.process_col_description(input, output)

    description_index = find_col(output,"description")
    kp_kr_index = find_col(input,"KPI / KR")
    description_indicador_index = find_col(input,"Descripción Indicador")
    i_index = find_col(input,"NC = 0%")
    j_index = find_col(input,"NC = 50%")
    k_index = find_col(input,"NC = 100%")
    l_index = find_col(input,"NC = 125%")
    m_index = find_col(input,"NC = 150%")
    rows = input&.sheet_data&.rows&.size
    i=1
    while i<rows.to_i
      kp_kr_val = input[i][kp_kr_index]&.value.to_s
      description_indicador_val = input[i][description_indicador_index]&.value.to_s
      i_val = input[i][i_index]&.value.to_s
      j_val = input[i][j_index]&.value.to_s
      k_val = input[i][k_index]&.value.to_s
      l_val = input[i][l_index]&.value.to_s
      m_val = input[i][m_index]&.value.to_s
      output_str=""
      if kp_kr_val.downcase == "kpi"
        if description_indicador_val.empty? || description_indicador_val == " "
          output_str = "<b>NC = 0%:</b>#{i_val}<br><b>NC = 50%: </b>#{j_val}<br><b>NC = 100% (Meta): </b>#{k_val}<br><b>NC = 125%: </b>#{l_val}<br><b>NC = 150%: </b>#{m_val}"
        else
          output_str = "#{description_indicador_val}<br><b>NC = 0%:</b> #{i_val}<br><b>NC = 50%: </b>#{j_val}<br><b>NC = 100% (Meta): </b>#{k_val}<br><b>NC = 125%: </b>#{l_val}<br><b>NC = 150%: </b>#{m_val}"
        end
      end
      
      if kp_kr_val.downcase == "kr"
        if description_indicador_val.empty? || description_indicador_val == " "
          output_str = "<b>NC = 100% (Meta):</b> #{k_val}"
        else
          output_str = "#{description_indicador_val}<br><b>NC = 100% (Meta):</b> #{k_val}"
        end
      end



      print(output,i,description_index,output_str)
      i+=1
    end
    

  end
  
  # rules for ponderation column
  private_class_method def self.process_col_ponderation(input, output)

    peso_final_index = find_col(input,"peso final")
    ponderation_index = find_col(output,"ponderation")
   
    rows = input&.sheet_data&.rows&.size
    i=1
    while i<rows.to_i
      peso_final_val= input[i][peso_final_index]&.value.to_i

     
        print(output,i,ponderation_index,peso_final_val)
      i+=1
    end
  end



  #--------------------step 2 file 1----------------

  # to add a header at end of sheet
  private_class_method def self.add_col_at_end(file,name)
    i=file[0].size
    file&.add_cell(0,i,name)
    return i
  end

  #to find a group code in rankmi tab for a combination (identifier, evaluator_identifier, tipo)
  private_class_method def self.find_groupcode(file, row, find_in)
    identifier_val = file[row][find_col(file,"identifier")]&.value
    identifier_index_rankmi = find_col(find_in,"identifier")
    evaluator_identifier_val =  file[row][find_col(file,"evaluator_identifier")]&.value
    evaluator_identifier_index_rankmi = find_col(find_in,"evaluator_identifier")
    tipo_val =  file[row][find_col(file,"tipo")]&.value
    tipo_index_rankmi = find_col(find_in,"tipo")
    groupcode_index_rankmi = find_col(find_in,"groupcode")
    groupcode = nil
    find_in.each_with_index do |row, i|
      
      if identifier_val==row[identifier_index_rankmi]&.value && evaluator_identifier_val ==row[evaluator_identifier_index_rankmi]&.value && tipo_val == row[tipo_index_rankmi]&.value
        groupcode=row[groupcode_index_rankmi]&.value
        break
      end
    end
    return groupcode
  end


  #to insert a header at start and shift all columns right
  private_class_method def self.insert_a_col_at_start(input,name)
    input.each_with_index do |row, i|
      size = row.size
      while size>0
        input.add_cell(i,size,row[size-1]&.value)
        size-=1
      end
      input.add_cell(i,size,i==0 ? name : "")
    end

  end


  # find a row in sheet in a column with header and a value to search in and to skip the column , row having g in dimension 
  private_class_method def self.find_row_with_dim_g(file,column_name,value,curr_col)
    i=curr_col
    found = nil
    column_index = find_col(file, column_name)
    dim_index = find_col(file, "dimension_type")
    return nil if !column_index
    while i < file.sheet_data.rows.size 
        break if !column_index
        if file[i][column_index]&.value&.downcase == value&.downcase && curr_col!=i && file[i][dim_index]&.value&.downcase == 'g'
            found=i
            break
        end
        i+=1
    end
    return found
  end

  # assign a groupcode to a row in es from rankmi
  private_class_method def self.process_the_identifier_code(input,rankmi)
    groupcode_index = find_col(input,"groupcode")
    identifier_code_index_rankmi = find_col(rankmi,"identifier_code")
    identifier_code_index = find_col(input,"identifier_code")

    used_identifier_codes = []

    input.each_with_index do |row, i|
      next if i==0
      groupcode_val = row[groupcode_index]&.value
      if groupcode_val=="N/A"
        input.add_cell(i,identifier_code_index, "N/A")
        next
      end

      found_groupcode_rankmi_index = find_row_with_dim_g(rankmi, "groupcode",groupcode_val,-1)
      if found_groupcode_rankmi_index
        found_identifier_code_rankmi_val = rankmi[found_groupcode_rankmi_index][identifier_code_index_rankmi]&.value
      end
      
      while !found_groupcode_rankmi_index.nil? && used_identifier_codes.select{|v| v[:identifier_code]==found_identifier_code_rankmi_val}.size>0  
        found_groupcode_rankmi_index = find_row_with_dim_g(rankmi, "groupcode",groupcode_val,found_groupcode_rankmi_index)
        if found_groupcode_rankmi_index
          found_identifier_code_rankmi_val = rankmi[found_groupcode_rankmi_index][identifier_code_index_rankmi]&.value
        end
      end

      if !found_groupcode_rankmi_index.nil?
        input.add_cell(i,identifier_code_index, found_identifier_code_rankmi_val)
        used_identifier_codes<<{"identifier_code": found_identifier_code_rankmi_val,"groupcode": groupcode_val}
      else
        last = used_identifier_codes.select{|v| v[:groupcode]==groupcode_val}.map{|v| v[:identifier_code]}.sort.last.to_s
        if(!last.empty?)
          new_val = last.split('-')
          new_val[new_val.size-1] = new_val.last.to_i+1
          new_val = new_val.join('-')

          input.add_cell(i,identifier_code_index, new_val)
          used_identifier_codes<<{"identifier_code": new_val,"groupcode": groupcode_val}
        end
      end
      
     
    end
  end

  # copy all sheets in a file and return a new one
  private_class_method def self.duplicate(input)
    output = RubyXL::Workbook.new
    j = 0
    while j < input.worksheets.size
      output.add_worksheet() if j>0
      input[j].each_with_index do |row, i|
        k=0
        while k < row.size
          output[j].add_cell(i,k,row[k]&.value)
          k+=1
        end
      end
      j+=1
    end
    return output
  end


  #-------------2nd file in step 2

  # copy a row at a index in file 
  private_class_method def self.copy_row(row,output,index)
    i=0
    while i< row.size
      output.add_cell(index, i , row[i]&.value)
      i+=1
    end
  end
  
  # delete a col and shift all columns to left
  private_class_method def self.delete_col(file,index)
    file.each_with_index do |row,i|
      k=index
      while k < row.size
        cell = row[k]
        next_cell = row[k+1]
        cell.change_contents(next_cell&.value)
        k+=1
      end
    end
  end



  #------------------ step 2 file 3 and 4-------------
  

  # get object having keys as unique values in a given column and there value to be their frequencies
  private_class_method def self.get_unique_and_freq_from_col(file,name)
    groupcode_index = find_col(file, name)
    dim_index = find_col(file, "dimension_type")
    obj = {}
    size = file.sheet_data.rows.size
    i = 1
    while i < size
      if file[i][dim_index]&.value == "g"
        cell = file[i][groupcode_index]
        if obj[cell&.value].nil?
          obj[cell&.value]=1
        else
          obj[cell&.value]+=1
        end
      end
      i+=1
    end
    # arr=[]
    # obj.keys.each do |val|
    #   arr<< {frequency: obj[val], groupcode: val }
    # end

    return obj
  end

  # create a file to make more frequency value than rankmi

  private_class_method def self.add_records_with_more_freq(input,output,val)
    groupcode_index = find_col(input, "groupcode")
    indentifier_index = find_col(input, "identifier")
    dim_index = find_col(input, "dimension_type")
    indentifier_code_index = find_col(input, "identifier_code")
    add_header_to_output(output,["identifier_code",	"identifier",	"groupcode"])

    input.each_with_index do |row,i|
      rows=output.sheet_data.rows.size
      if row[dim_index]&.value == "g" && row[groupcode_index]&.value == val
        output.add_cell(rows, 0, row[indentifier_code_index]&.value)
        output.add_cell(rows, 1, row[indentifier_index]&.value)
        output.add_cell(rows, 2, val)
      end
    end
    

  end

  end
end
