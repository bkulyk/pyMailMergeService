load 'deploy' if respond_to?(:namespace) # cap2 differentiator
require 'capistrano_colors'

# Standard configuration
#set :user, "ubuntu"
#set :password, "password"
set :application, "pyMailMergeService"

# I like to deploy the code in /var/apps
# and then link it to the webserver directory
set :deploy_to, "/home/ubuntu/#{application}"

set :repository, "git://github.com/bkulyk/pyMailMergeService.git"
set :scm, :git
set :branch, "pyMailMerge.2.0"
set :repository_cache, 'git_cache'
set :deploy_via, :remote_cache
set :scm_verbose, true
set :keep_releases, 10

role :app, "mms", :primary => true

namespace :deploy do
	
	desc <<-DESC
		Restart the python application and openoffice
	DESC
	task :restart, :roles=>:app do
		run "#{latest_release}/restartOOandPMMS.sh"
	end
	task :finalize_update, :except => { :no_release => true  } do
		run "chmod -R g+w #{latest_release}" if fetch( :group_writable, true )
	end
end
