# frozen_string_literal: true

require "microsoft/graph"
require "vcr"
require "awesome_print"

def use_cassette(name)
  before(:each) do
    VCR.insert_cassette name
  end

  after(:each) do
    VCR.eject_cassette name
  end
end

VCR.configure do |config|
  config.cassette_library_dir = "spec/cassettes"
  # config.default_cassette_options = { record: :new_episodes }
  config.hook_into :webmock
end

RSpec.configure do |config|
  # Enable flags like --only-failures and --next-failure
  config.example_status_persistence_file_path = ".rspec_status"

  # Disable RSpec exposing methods globally on `Module` and `main`
  config.disable_monkey_patching!

  config.expect_with :rspec do |c|
    c.syntax = :expect
  end
end
