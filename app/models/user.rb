class User < ApplicationRecord
  # Include default devise modules. Others available are:
  # :confirmable, :lockable, :timeoutable, :trackable and :omniauthable
  devise :database_authenticatable, :registerable,
         :recoverable, :rememberable, :validatable
  devise :omniauthable, omniauth_providers: [:google_oauth2]

  def self.from_omniauth(auth)
    user = self.find_or_initialize_by(provider: auth.provider, uid: auth.uid)
    user.email = auth.info.email
    user.password ||= Devise.friendly_token[0, 20]
    user.refresh_token = auth.credentials.refresh_token
    user.access_token = auth.credentials.token
    user.expires_at = auth.credentials.expires_at
    user.save
    user
  end
end
