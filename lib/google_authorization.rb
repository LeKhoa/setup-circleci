require "google/api_client/client_secrets.rb"

class GoogleAuthorization

  def initialize(user)
    @user = user
  end

  def authorize
    secrets = Google::APIClient::ClientSecrets.new({
      web: {
        access_token: @user.access_token,
        refresh_token: @user.refresh_token,
        client_id: Rails.application.credentials.google[:client_id],
        client_secret: Rails.application.credentials.google[:client_secret],
        expires_at: @user.expires_at,
      }
    })

    authorization = secrets.to_authorization
    expires_soon_at = DateTime.current + 10.minutes

    if Time.at(@user.expires_at) <= expires_soon_at
      authorization.fetch_access_token!
      @user.access_token = authorization.access_token
      @user.expires_at = authorization.expires_at
      @user.save
    end

    authorization
  end
end
