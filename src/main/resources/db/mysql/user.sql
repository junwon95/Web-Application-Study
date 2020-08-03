CREATE DATABASE IF NOT EXISTS petclinic;

ALTER DATABASE petclinic
  DEFAULT CHARACTER SET utf8
  DEFAULT COLLATE utf8_general_ci;

--****** might be a problem when deployed to Heroku (probably uses different username)
GRANT ALL PRIVILEGES ON petclinic.* TO 'petclinic';
