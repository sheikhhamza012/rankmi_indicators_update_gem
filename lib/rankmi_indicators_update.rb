# require "rankmi_indicators_update/version"
require_relative "rankmi_indicators_update/helpers"
require 'rubyXL'
# require 'byebug'
require 'rubyXL/convenience_methods'

module RankmiIndicatorsUpdate
  class Error < StandardError; end
  class Parse < Helper
    
    #--------step 1--------
  
    def self.transform_to_rankmi_template(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = RubyXL::Workbook.new

      add_header_to_output(output[0],%w[identifier evaluator_identifier dimension_type tipo groupcode title description min meta max evaluation_type due_date unit ponderation expected_value score evaluable editable show_current_process editable_ponderation replicator_tag])
      copy_column(input[0],"Matrícula",output[0],"identifier")
      copy_column(input[0],"evaluator_identifier",output[0],"evaluator_identifier")
      copy_value_in_col(output[0],"dimension_type","g")
      copy_value_in_col(output[0],"show_current_process","x")
      copy_value_in_col(output[0],"evaluation_type","prc_simpl")
      copy_value_in_col(output[0],"unit","%")
      copy_column(input[0],"KPI / KR",output[0],"tipo","s")
      copy_column(input[0],"Nombre de indicador",output[0],"title")
      process_col_description(input[0],output[0])
      process_col_ponderation(input[0],output[0])
      copy_column(input[0],"NC = 100%",output[0],"expected_value")
      copy_column(input[0],"Resultado",output[0],"score")
      
      # output.write("Metas BCP.xlsx")
      return output
    end


    #--------step 2--------

    def self.first_output_of_step_two(file_path)
      input = RubyXL::Parser.parse(file_path)
      input = duplicate(input)
      index_of_tipo = add_col_at_end(input[1],"tipo")
      index_of_title = find_col(input[1],"title")
      # index_of_groupcode = find_col(input[1],"groupcode")
      index_of_dimension_type = find_col(input[1],"dimension_type")
      title = ""
      input[1].each_with_index do |row, i|
        next if i==0
        val_of_title = row[index_of_title]&.value
        # val_of_groupcode = input[0][i][index_of_groupcode]&.value
        
        val_of_dimension_type = row[index_of_dimension_type]&.value
        if val_of_dimension_type.downcase == 'o'.downcase
          title=val_of_title
        end
        print(input[1],i,index_of_tipo,title)
      end
      

      groupcode_index = find_col(input[0],"groupcode")
      input[0].each_with_index do |row, i|
        next if i==0
        groupcode = find_groupcode(input[0],i,input[1])
        if !groupcode
          groupcode= "N/A"
        end
        print(input[0],i,groupcode_index,groupcode)
      end

      insert_a_col_at_start(input[0], "identifier_code")
      process_the_identifier_code(input[0],input[1])
      
      # input.write("Metas BCP con info rankmi y tipo.xlsx")
      return input
  
    end

    # ------------step 2 file 2 ---------

    def self.separate_records(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = RubyXL::Workbook.new
      groupcode_index = find_col(input[0],"groupcode")
      copy_row(input[0][0], output[0],0)
      output_row = 1
      input[0].each_with_index do |row,i|
        groupcode_val = row[groupcode_index]&.value
        next if groupcode_val != "N/A"
        copy_row(row,output[0],output_row)
        output_row+=1
      end
      
      delete_col(output[0],find_col(output[0],"tipo"))
      delete_col(output[0],find_col(output[0],"identifier_code"))
      # output.write('Metas a crear.xlsx')
      return output
    end

  #------------------ step 2 file 3 and 4-------------

    def self.define_goals_to_eliminate_and_update(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = RubyXL::Workbook.new
      metas_a_borar_output = RubyXL::Workbook.new
  
      add_header_to_output(output[0],["#g",	"Pestaña \"es\"",	"Pestaña \"rankmi\"",	"Acción"])
      unique_groupcode_es = get_unique_and_freq_from_col(input[0],"groupcode")
      unique_groupcode_rankmi = get_unique_and_freq_from_col(input[1],"groupcode")
      unique_groupcode_es.keys.each_with_index do |val,i|
        next if val=="N/A"
        output[0].add_cell(i+1, 0 , val)
        output[0].add_cell(i+1, 1 , unique_groupcode_es[val])
        output[0].add_cell(i+1, 2 , unique_groupcode_rankmi[val])
        if unique_groupcode_es[val].to_i >= unique_groupcode_rankmi[val].to_i
          info="Actualizar"
        elsif unique_groupcode_es[val].to_i < unique_groupcode_rankmi[val].to_i
          add_records_with_more_freq(input[1],metas_a_borar_output[0],val)
          info="Eliminar en Rankmi"
        end
        if unique_groupcode_es[val].to_s.empty?
          info = "N/A"
        end
        output[0].add_cell(i+1, 3 , info)
      end
      # output.write('Análisis accione.xlsx')
      # metas_a_borar_output.write('Metas a borrar.xlsx')
      return {analisi: output, borrar: metas_a_borar_output}
    end


  #------------------- step 2 file 5-------------------

    def self.excel_with_goals_to_update(file_path)
      input = RubyXL::Parser.parse(file_path)
      output = RubyXL::Workbook.new
      
      groupcode_index = find_col(input[0],"groupcode")
      input[0].each_with_index do |row, i |
        next if row[groupcode_index]&.value == "N/A"
        copy_row(row, output[0], output[0].sheet_data.rows.size)
        
      end
      
      tipo_index = find_col(output[0],"tipo")
      codigo_index = find_col(output[0],"codigo apoyo")
      if tipo_index
        delete_col(output[0],tipo_index)
      end
      if codigo_index
        delete_col(output[0],codigo_index)
      end
      # output.write('Metas a actualizar.xlsx')
      return output
    end
  end
end
