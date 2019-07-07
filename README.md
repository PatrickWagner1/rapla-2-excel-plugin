# rapla-2-excel-plugin
You can integrate the plugin into the rapla client. You can use it analogous to the csv export function.

# Requirements
Java SDK 8
Eclipse (tested with Eclipse Photon and higher versions)

# Checkout projects from Git
1. Got to `Checkout projects from Git` in Eclipse (you will find it by clicking on Help > Welcome)
2. Select `Clone URI`
3. Enter `https://github.com/rapla/rapla` as URI
4. Select the 1_8 branch
5. Select import existing eclipse projects
6. Repeat step 1-5 for the URI `https://github.com/patrickwagner1/rapla-2-excel-plugin` and the master branch

# Setup the Run Configurations
1. Right click on `org.rapla.bootstrap.RaplaStandaloneLoader.java` in the src directory of the rapla project
2. Click on run As > Java Application
3. Close rapla
4. Go to the `Run Configurations...`
5. Select `Java Application > RaplastandaloneLoader > Classpath > User Entries`
6. Click on `Add Projects...`
7. Select `rapla-2-excel-plugin`
8. Click `OK`
9. Click on `Add Jars...`
10. Select all Jars in rapla-2-exel-plugin/poi-4.1.0 and rapla-2-excel-plugin/jollyday-0.5.1.jar
11. Click `OK`
12. Click `Apply`
13. Close the Window of the Run Configurations

# Start Rapla with the Excel Plugin
You can start Rapla by clicking on Run in Eclipse with the RaplaStandaloneLoader Run Configurations.

# Start the Standalone Version of the Converter
1. Right Click on `semesterTimeTable.excel.standalone.Standalone.java` in the src directory of the rapla-2-excel-plugin Project
2. Click on run As > Java Application
