namespace :keywords do
  task automate: :environment do
    KeywordAutomator.call
  end
end
