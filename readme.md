<<<<<<< HEAD
# Spring PetClinic Sample Application
Base sample code from: https://github.com/spring-projects/spring-petclinic.git

## Running the application locally
Petclinic is a [Spring Boot](https://spring.io/guides/gs/spring-boot) application built using [Maven](https://spring.io/guides/gs/maven/). You can build a jar file and run it from the command line:


```
git clone https://github.com/junwon95/Web-Application-Study.git
cd Web-Application-Study
./mvnw package
java -jar target/*.jar
```

You can then access the application here: http://localhost:8080/


## Database configuration

The default setup is set to MySql.

create user 'petclinic' on localhost(8080).
then run src/main/resources/db/mysql/user.sql on MySql to initialize. (Once only) 

You could start MySql locally with whatever installer works for your OS

## Mail configuration

On src/main/resources/application.yml
fill in your email account and password.

## Working with Petclinic in your IDE

### Prerequisites
The following items should be installed in your system:
* Java 8 or newer.
* git command line tool (https://help.github.com/articles/set-up-git)
* Your preferred IDE 
  * Eclipse with the m2e plugin. Note: when m2e is available, there is an m2 icon in `Help -> About` dialog. If m2e is
  not there, just follow the install process here: https://www.eclipse.org/m2e/
  * [Spring Tools Suite](https://spring.io/tools) (STS)
  * IntelliJ IDEA *** recommended (the IDE I used to develop the application)
  * [VS Code](https://code.visualstudio.com)

### Steps:

1) On the command line
    ```
    git clone https://github.com/junwon95/Web-Application-Study.git
    ```
2) Inside Eclipse or STS
    ```
    File -> Import -> Maven -> Existing Maven project
    ```

    Then either build on the command line `./mvnw generate-resources` or using the Eclipse launcher (right click on project and `Run As -> Maven install`) to generate the css. Run the application main method by right clicking on it and choosing `Run As -> Java Application`.

3) Inside IntelliJ IDEA
    In the main menu, choose `File -> Open` and select the Petclinic [pom.xml](pom.xml). Click on the `Open` button.

    CSS files are generated from the Maven build. You can either build them on the command line `./mvnw generate-resources` or right click on the `spring-petclinic` project then `Maven -> Generates sources and Update Folders`.

    A run configuration named `PetClinicApplication` should have been created for you if you're using a recent Ultimate version. Otherwise, run the application by right clicking on the `PetClinicApplication` main class and choosing `Run 'PetClinicApplication'`.

4) Navigate to Petclinic

    Visit [http://localhost:8080](http://localhost:8080) in your browser.
=======
# Spring PetClinic Sample Application
Base sample code from: https://github.com/spring-projects/spring-petclinic.git

## Running the application locally
Petclinic is a [Spring Boot](https://spring.io/guides/gs/spring-boot) application built using [Maven](https://spring.io/guides/gs/maven/). You can build a jar file and run it from the command line:


```
git clone https://github.com/junwon95/Web-Application-Study.git
cd Web-Application-Study
./mvnw package
java -jar target/*.jar
```

You can then access the application here: http://localhost:8080/


## Database configuration

The default setup is set to MySql.

create user 'petclinic' on localhost(8080).
then run src/main/resources/user.sql on MySql to initialize. (Once only) 

You could start MySql locally with whatever installer works for your OS

## Working with Petclinic in your IDE

### Prerequisites
The following items should be installed in your system:
* Java 8 or newer.
* git command line tool (https://help.github.com/articles/set-up-git)
* Your preferred IDE 
  * Eclipse with the m2e plugin. Note: when m2e is available, there is an m2 icon in `Help -> About` dialog. If m2e is
  not there, just follow the install process here: https://www.eclipse.org/m2e/
  * [Spring Tools Suite](https://spring.io/tools) (STS)
  * IntelliJ IDEA *** recommended (the IDE I used to develop the application)
  * [VS Code](https://code.visualstudio.com)

### Steps:

1) On the command line
    ```
    git clone https://github.com/junwon95/Web-Application-Study.git
    ```
2) Inside Eclipse or STS
    ```
    File -> Import -> Maven -> Existing Maven project
    ```

    Then either build on the command line `./mvnw generate-resources` or using the Eclipse launcher (right click on project and `Run As -> Maven install`) to generate the css. Run the application main method by right clicking on it and choosing `Run As -> Java Application`.

3) Inside IntelliJ IDEA
    In the main menu, choose `File -> Open` and select the Petclinic [pom.xml](pom.xml). Click on the `Open` button.

    CSS files are generated from the Maven build. You can either build them on the command line `./mvnw generate-resources` or right click on the `spring-petclinic` project then `Maven -> Generates sources and Update Folders`.

    A run configuration named `PetClinicApplication` should have been created for you if you're using a recent Ultimate version. Otherwise, run the application by right clicking on the `PetClinicApplication` main class and choosing `Run 'PetClinicApplication'`.

4) Navigate to Petclinic

    Visit [http://localhost:8080](http://localhost:8080) in your browser.
>>>>>>> 26cc5b0aa6e81eece08cca3e24047ed31c8aa1ea
