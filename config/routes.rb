Rails.application.routes.draw do
  devise_for :users, controllers: { omniauth_callbacks: "users/omniauth_callbacks" }
  root "welcome#index"

  namespace :users do
    root "home#index"
  end
end
