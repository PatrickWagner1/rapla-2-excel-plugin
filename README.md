# rapla-2-excel-plugin
You can integrate the plugin into the rapla client. You can use it analogous to the csv export function.

# Setup
> You should use Java SDK 1.8

You need two projects in Eclipse, rapla itself and the excel plugin.

## Create rapla project
> You should take the 1_8 branch of rapla.

To create the rapla project, follow the steps of [BuildingRaplaWithEclipse](https://github.com/rapla/rapla/wiki/BuildingRaplaWithEclipse).

## Create excel plugin
To create the rapla to excel plugin, create a project from git in eclipse with [this repository](https://github.com/patrickwagner1/rapla-2-excel-plugin).

## Link the two projects.
1. Rightclick on `org.rapla.bootstrap.RaplaStandaloneLoader` in the rapla project and select Run As -> Java Application to create a run entry.
2. Close Rapla and rightclick again on `org.rapla.bootstrap.RaplaStandaloneLoader` and select Run As -> Run Configuration...
3. Add the rapla-2-excel-plugin project to the classpath.
4. Rightclick on the rapla-2-excel-plugin and select properties
5. Add the rapla project to the projects
6. Run the created run configurations for the `RaplaStandaloneLoader`
