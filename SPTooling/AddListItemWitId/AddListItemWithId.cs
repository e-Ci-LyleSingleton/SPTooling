using Microsoft.SharePoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Coliban.SharePoint.Tooling
{
    class AddListItemWithId
    {
        protected static int itemId;

        protected static string siteUrl;
        protected static string webRelativeUrl;
        protected static string listName;
        protected static string propMappingFile;
        
        protected static SPSite site;
        protected static SPWeb web;
        protected static SPList list;


        static void Main(string[] args)
        {
            itemId = 0;
            siteUrl = string.Empty;
            webRelativeUrl = string.Empty;
            listName = string.Empty;
            propMappingFile = string.Empty;

            for (int i = 0; i < args.Length; i++)
            {
                string argumentString = args[i];

                string argumentValue = string.Empty;

                if (i + 1 < args.Length)
                {
                    argumentValue = args[i + 1].Trim();
                }

                switch (argumentString)
                {
                    case "-itemId":
                        {
                            int.TryParse(argumentValue, out itemId);
                            i++;
                            break;
                        }
                    case "-siteUrl":
                        {
                            siteUrl = argumentValue;
                            i++;
                            break;
                        }
                    case "-webUrl":
                        {
                            webRelativeUrl = argumentValue;
                            i++;
                            break;
                        }
                    case "-listName":
                        {
                            listName = argumentValue;
                            i++;
                            break;
                        }
                    case "-propMapping":
                        {
                            propMappingFile = argumentValue;
                            i++;
                            break;
                        }
                    default:
                        Console.WriteLine("Unknown argument: " + argumentString);
                        return;
                }
            }


            if (itemId == 0)
            {
                Console.WriteLine("you must specify a valid item Id with -itemId <Id>");
                return;
            }
            else if (siteUrl == string.Empty)
            {
                Console.WriteLine("you must specify a site collection url with -siteUrl <SiteUrl>");
                return;
            }
            else if (webRelativeUrl == string.Empty)
            {
                Console.WriteLine("you must specify a site collection relative site url with -webUrl <WebUrl>");
                return;
            }
            else if (listName == string.Empty)
            {
                Console.WriteLine("you must specify a SharePoint list name with -listName <ListName>");
                return;
            }
            else if (propMappingFile == string.Empty)
            {
                Console.WriteLine("you must specify a property mapping file with -propMapping <FilePath>");
                return;
            }

            try
            {

                Console.WriteLine("Connecting to site collection at '" + siteUrl + "'");
                site = new SPSite(siteUrl);

                web = site.OpenWeb(webRelativeUrl);
                Console.WriteLine("Opened site '" + web.Title + "'");

                list = web.Lists[listName];
                Console.WriteLine("Opened list '" + list.Title + "'");
                
                SPQuery existingItemQuery = new SPQuery();
                existingItemQuery.ViewXml = @"
<View>
    <Query>
        <Where>
            <Eq>
                <FieldRef Name='ID' />
                <Value Type='Integer'>" + itemId.ToString() + @"</Value>
            </Eq>
        </Where>
    </Query>
    <ViewFields>
      <FieldRef Name='Title'/>
    </ViewFields>
</View>";

                SPListItemCollection existingItems = list.GetItems(existingItemQuery);
                if( existingItems.Count != 0 )
                {
                    throw new ArgumentOutOfRangeException("Item id '" + existingItems[0][SPBuiltInFieldId.Title]+"' is already in use.");
                }

                // The ID field is readonly field, set to to read and write mode.
                list.Fields[SPBuiltInFieldId.ID].ReadOnlyField = false;
                list.Update();

                JsonSerializerSettings settings = new JsonSerializerSettings { Converters = new JsonConverter[] { new Common.JsonGenericDictionaryOrArrayConverter() } };

                Dictionary<string, string> propMaps = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(propMappingFile), settings);

                SPListItem item = list.AddItem();

                foreach (string key in propMaps.Keys)
                {
                    item[key] = propMaps[key];
                }

                item[SPBuiltInFieldId.ID] = itemId;
                item.Update();
            }
            catch (System.Exception e)
            {
                System.Console.WriteLine(e.Message);
            }
            finally
            {
                if( list != null )
                {
                    if (list.Fields[SPBuiltInFieldId.ID].ReadOnlyField != true)
                    {
                        list.Fields[SPBuiltInFieldId.ID].ReadOnlyField = true;
                        list.Update();
                    }
                }
                if( site != null )
                {
                    site.Dispose();
                }
                if( web != null )
                {
                    web.Dispose();
                }
            }
        }
    }
}
