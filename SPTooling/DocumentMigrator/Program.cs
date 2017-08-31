using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentMigrator
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext sourceCtx = new ClientContext(Properties.Resources.sourceWebUrl);
            ClientContext destCtx = new ClientContext(Properties.Resources.destWebUrl);

            List sourceList = sourceCtx.Web.Lists.GetByTitle(Properties.Resources.sourceDocLibName);
            List destList = destCtx.Web.Lists.GetByTitle(Properties.Resources.destDocLibName);

            View migratingItemsView = sourceList.Views.GetById(new Guid(Properties.Resources.sourceViewId));
            
            sourceCtx.Load(migratingItemsView);
            sourceCtx.ExecuteQuery();

            CamlQuery migratingItemsCaml = new CamlQuery();
            migratingItemsCaml.ViewXml = migratingItemsView.HtmlSchemaXml;

            ListItemCollection migratingItemsList = sourceList.GetItems(migratingItemsCaml);

            sourceCtx.Load(migratingItemsList);
            sourceCtx.ExecuteQuery();

            IEnumerator<ListItem> itemIter = migratingItemsList.GetEnumerator();

            while(itemIter.MoveNext())
            {
                ListItem itemToMigrate = itemIter.Current;
                if ( itemToMigrate.FileSystemObjectType == FileSystemObjectType.Folder )
                {
                    string libRelFolderUrl = GetNormalisedFileRef(itemToMigrate);
                    string folderName = itemToMigrate.FieldValues["FileLeafRef"] as string;
                    DateTime modifiedDate = (DateTime)itemToMigrate.FieldValues["Modified"];
                    Console.WriteLine(modifiedDate + " Folder: " + libRelFolderUrl);
                    // Check if folder exists
                }
                else if( itemToMigrate.FileSystemObjectType == FileSystemObjectType.File )
                {
                    string libRelFolderUrl = GetNormalisedFileRef(itemToMigrate);
                    string fileName = itemToMigrate.FieldValues["FileLeafRef"] as string;
                    int fileNameIndex = libRelFolderUrl.LastIndexOf(fileName);
                    string sourceFolder = libRelFolderUrl.Remove(fileNameIndex-1, fileName.Length+1);
                    string author = (itemToMigrate.FieldValues["Author"] as FieldUserValue).LookupValue;
                    string modifier = (itemToMigrate.FieldValues["Editor"] as FieldUserValue).LookupValue;
                    DateTime modified = (DateTime)itemToMigrate.FieldValues["Modified"];
                    DateTime created = (DateTime)itemToMigrate.FieldValues["Created"];

                    Console.WriteLine(modified + " \"" + fileName + "\" in \"" + sourceFolder + "\"");

                    // check if file already exists
                }
            }
            Console.ReadLine();
        }

        static string GetNormalisedFileRef(ListItem item)
        {
            string result = item.FieldValues["FileRef"] as string;

            if( !item.ParentList.RootFolder.IsObjectPropertyInstantiated("ServerRelativeUrl") )
            {
                item.Context.Load(item.ParentList.RootFolder);
                item.Context.ExecuteQuery();
            }

            string libraryRootUrl = item.ParentList.RootFolder.ServerRelativeUrl;
            int libraryRootUrlIndex = result.IndexOf(libraryRootUrl);

            if (libraryRootUrlIndex == 0 )
            {
                result = result.Substring(libraryRootUrl.Length, result.Length - libraryRootUrl.Length);
            }
            else
            {
                throw new ArgumentException("fileRef from library", "list");
            }
            return result;
        }
    }
}


/*
  Public Shared Sub UploadStreamToSharepoint(ByVal oMemoryStream As MemoryStream, _
                                            ByVal fileName As String, ByVal sLibraryName As String, ByVal sRelativePathToLibrary As String, _
                                            Optional ByVal sDrawingNumber As String = "", Optional ByVal fromException As Boolean = False, Optional hasWatermark As Boolean = False)

        On Error Resume Next

        Dim siteUrl As String = c.gV(Common.MyApp.Setting("SiteURL")).Trim()

        Using context = New ClientContext(siteUrl)

            Dim securePassword As New SecureString()
            For Each c As Char In Common.Crypto.DES.DeCrypt(Common.MyApp.Setting("globalPassword")).ToCharArray()
                securePassword.AppendChar(c)
            Next
            context.AuthenticationMode = ClientAuthenticationMode.[Default]
            context.Credentials = New SharePointOnlineCredentials(Common.MyApp.Setting("globalUsername"), securePassword)
            

            Dim web As ClientOM.Web = context.Web
            Dim list As ClientOM.List = web.Lists.GetByTitle(sLibraryName)
            context.ExecuteQuery()


            oMemoryStream.Position = 0

            Dim fileServerRelativeUrl As String = sRelativePathToLibrary & "/" & Common.c.gV(fileName).Trim()

            Using oMemoryStream

                ClientOM.File.SaveBinaryDirect(context, fileServerRelativeUrl, oMemoryStream, True)

                context.ExecuteQuery()

            End Using

        End Using

        On Error GoTo 0

    End Sub

     
     */
