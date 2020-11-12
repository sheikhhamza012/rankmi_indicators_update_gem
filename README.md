# RankmiIndicatorsUpdate

Welcome to your new gem! In this directory, you'll find the files you need to be able to package up your Ruby library into a gem. Put your Ruby code in the file `lib/rankmi_indicators_update`. To experiment with that code, run `bin/console` for an interactive prompt.

TODO: Delete this and the text above, and describe your gem

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'rankmi_indicators_update', git: 'https://github.com/sheikhhamza012/rankmi_indicators_update_gem.git'
```

And then execute:

    $ bundle


## Usage

- ` RankmiIndicatorsUpdate::Parse.transform_to_rankmi_template('excel_files/excel.xlsx')` will create file `Metas BCP.xlsx`
- ` RankmiIndicatorsUpdate::Parse.first_output_of_step_two('excel_files/excel-2.xlsx')` will create file ` Metas BCP con info rankmi y tipo.xlsx `
- ` RankmiIndicatorsUpdate::Parse.separate_records('Metas BCP con info rankmi y tipo.xlsx')` will create file `Metas a crear.xlsx`
- ` RankmiIndicatorsUpdate::Parse.define_goals_to_eliminate_and_update('Metas BCP con info rankmi y tipo.xlsx')` will create files ` Análisis accione.xlsx` and `Metas a borrar.xlsx`
- ` RankmiIndicatorsUpdate::Parse.excel_with_goals_to_update('Metas BCP con info rankmi y tipo.xlsx')` will create file `Metas a actualizar.xlsx`

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/rankmi_indicators_update. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the RankmiIndicatorsUpdate project’s codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/[USERNAME]/rankmi_indicators_update/blob/master/CODE_OF_CONDUCT.md).
