require "open-uri"
require "google/apis/drive_v3"
require "google/apis/sheets_v4"
require "google/api_client/client_secrets.rb"
require "screenshot"
require "slide"

class KeywordAutomator
  def self.call(options = {})
    user = User.find(options[:user_id])
    secrets = Google::APIClient::ClientSecrets.new({
      web: {
        access_token: user.access_token,
        refresh_token: user.refresh_token,
        client_id: Rails.application.credentials.google[:client_id],
        client_secret: Rails.application.credentials.google[:client_secret],
      }
    })

    sheets_service = ::Google::Apis::SheetsV4::SheetsService.new
    sheets_service.authorization = secrets.to_authorization
    keyword_file_id = options[:keyword_file_id]
    keywords_range = "A2:A1000"
    included_keywords_range = "B2:B1000"
    ranges = [keywords_range, included_keywords_range]
    result = sheets_service.batch_get_spreadsheet_values(keyword_file_id, ranges: ranges)
    keywords_value_range = result.value_ranges.detect { |vr| vr.range.include?(keywords_range) }
    keywords = keywords_value_range.values.flatten
    included_keywords_value_range = result.value_ranges.detect { |vr| vr.range.include?(included_keywords_range) }
    included_keywords = included_keywords_value_range.values.flatten
    Screenshot.new(keywords, included_keywords).call

    folder_id = options[:result_folder_id]
    filename = "brand_control_template"
    file_metadata = {
      name: filename,
      mime_type: "application/vnd.google-apps.presentation",
      parents: [folder_id],
    }

    drive_service = ::Google::Apis::DriveV3::DriveService.new
    drive_service.authorization = secrets.to_authorization
    template_file_id = "1VdkrDyjfm2hL3uAa0x6mHlt7PAbOcE06"
    template_url = "https://drive.google.com/uc?id=#{template_file_id}&export=download"
    template_file = open(template_url)
    file = drive_service.create_file(
      file_metadata,
      fields: "id",
      upload_source: template_file,
      content_type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    Slide.new(user, file.id).create_for_all_keywords
  end
end
