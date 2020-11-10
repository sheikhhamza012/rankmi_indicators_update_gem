require "rankmi_indicators_update/version"
require "rankmi_indicators_update/helpers"

module RankmiIndicatorsUpdate
  class Error < StandardError; end
  class Parse
    
    def self.transform_to_rankmi_template
      RankmiIndicatorsUpdate::Helper.say_hello
    end

  end
end
