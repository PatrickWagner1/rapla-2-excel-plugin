package org.rapla.plugin.export2excel;

import org.rapla.client.ClientServiceContainer;
import org.rapla.client.RaplaClientExtensionPoints;
import org.rapla.components.xmlbundle.I18nBundle;
import org.rapla.framework.Configuration;
import org.rapla.framework.PluginDescriptor;
import org.rapla.framework.TypedComponentRole;

/**
 * Class representing the export to excel plugin within Rapla.
 */
public class Export2ExcelPlugin implements PluginDescriptor<ClientServiceContainer> {

	/** A boolean that determines if the plugin is enabled by default. */
	public static final boolean ENABLE_BY_DEFAULT = true;

	/** The xml file containing the resources for the plugin. */
	public static final TypedComponentRole<I18nBundle> RESOURCE_FILE = new TypedComponentRole<I18nBundle>(
			Export2ExcelPlugin.class.getPackage().getName() + ".ExcelResources");

	public void provideServices(ClientServiceContainer container, Configuration config) {
		if (config.getAttributeAsBoolean("enabled", ENABLE_BY_DEFAULT)) {

			container.addResourceFile(RESOURCE_FILE);

			// container.addContainerProvidedComponent(RaplaClientExtensionPoints.USER_OPTION_PANEL_EXTENSION,
			// MyOption.class);

			container.addContainerProvidedComponent(RaplaClientExtensionPoints.EXPORT_MENU_EXTENSION_POINT,
					Export2ExcelMenu.class);

			// container.addContainerProvidedComponent(RaplaClientExtensionPoints.HELP_MENU_EXTENSION_POINT,
			// MyHelpMenuExtension.class);

		}
	}
}
