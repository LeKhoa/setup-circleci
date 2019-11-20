require "google/apis/drive_v3"

module Users
  class HomeController < BaseController
    before_action :authenticate_user!

    def index
      drive_service = ::Google::Apis::DriveV3::DriveService.new

      if params[:keyword_file_id]
        AutomateKeywordsJob.perform_later(
          user_id: current_user.id,
          keyword_file_id: params[:keyword_file_id],
          result_folder_id: params[:result_folder_id]
        )
        flash[:success] = "Automate successfully! Please check result in your drive in few minutes"
        redirect_to users_root_path
      end

      drive_service.authorization = GoogleAuthorization.new(current_user).authorize
      files = drive_service.list_files.files
      @spreadsheets = files.select { |f| f.mime_type == "application/vnd.google-apps.spreadsheet" }
      @folders = files.select { |f| f.mime_type == "application/vnd.google-apps.folder" }
    rescue
    end

  end
end
