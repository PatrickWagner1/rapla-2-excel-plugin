package org.rapla.plugin.export2excel;

import org.rapla.client.ClientServiceContainer;
import org.rapla.client.RaplaClientExtensionPoints;
import org.rapla.components.xmlbundle.I18nBundle;
import org.rapla.framework.Configuration;
import org.rapla.framework.PluginDescriptor;
import org.rapla.framework.TypedComponentRole;

public class Export2ExcelPlugin  implements PluginDescriptor<ClientServiceContainer>
{
	public static final boolean ENABLE_BY_DEFAULT = true;
	public static final TypedComponentRole<I18nBundle> RESOURCE_FILE = new TypedComponentRole<I18nBundle>(Export2ExcelPlugin.class.getPackage().getName() + ".ExcelResources");

	public void provideServices(ClientServiceContainer container, Configuration config) {
	  if ( !config.getAttributeAsBoolean("enabled", ENABLE_BY_DEFAULT) )
	     return;

	  container.addResourceFile(RESOURCE_FILE);
	  
	  //container.addContainerProvidedComponent( RaplaClientExtensionPoints.USER_OPTION_PANEL_EXTENSION, MyOption.class);

	  container.addContainerProvidedComponent( RaplaClientExtensionPoints.EXPORT_MENU_EXTENSION_POINT, Export2ExcelMenu.class);

	  //container.addContainerProvidedComponent( RaplaClientExtensionPoints.HELP_MENU_EXTENSION_POINT, MyHelpMenuExtension.class);

	}
	
}
