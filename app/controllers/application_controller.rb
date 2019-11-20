class ApplicationController < ActionController::Base
  protected

  def after_sign_in_path_for(resource)
    users_root_path
  end
end
