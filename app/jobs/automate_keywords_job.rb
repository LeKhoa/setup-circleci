class AutomateKeywordsJob < ApplicationJob
  queue_as :default
  sidekiq_options retry: false

  def perform(options)
    # Do something later
    KeywordAutomator.call(options)
  rescue
  end
end
