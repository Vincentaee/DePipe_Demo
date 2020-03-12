using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;
using Autodesk.Forge;
using Autodesk.Forge.Client;
using System.Diagnostics;
using Autodesk.Forge.Model;
using Newtonsoft.Json;
using System.IO;
using ExReaderConsole;

namespace cefsharpTest
{
    public partial class ForgeForm : Form
    {

        private ChromiumWebBrowser browser;
        string token = null;
        string doc_path = "";

        public ForgeForm(string [] args)
        {
            InitializeComponent();
            InitBrowser();
            doc_path = args[0];
        }

        private void ForgeForm_Load(object sender, EventArgs e)
        {
            var settings = new CefSettings();

            Cef.Initialize(settings);

        }

        public void InitBrowser()
        {



            //browser = new ChromiumWebBrowser("www.google.com");
            //browser = new ChromiumWebBrowser("https://developer.api.autodesk.com/authentication/v1/authorize?client_id=AYPBXmBpIzEBaGKUI55anxorccBBhpci&response_type=code&redirect_uri=http%3a%2f%2flocalhost%3a3000%2fapi%2fforge%2fcallback%2foauth&scope=data%3aread+data%3awrite+data%3acreate+bucket%3aread+bucket%3acreate");
            //browser = new ChromiumWebBrowser("http://localhost:3000/api/forge/callback/oauth?code=Volh1ip6xtLM2zUPJDctqK-TCXWHuQkRGqnHgRAy");
            //browser = new ChromiumWebBrowser("123");

            ThreeLeggedApi oauthree = new ThreeLeggedApi();
            string client_id = "AYPBXmBpIzEBaGKUI55anxorccBBhpci";
            string response_type = "token";
            string redirect_uri = "http://localhost:3000/api/forge/callback/oauth";
            /*
            string client_id = "4AaTAEzFFp54ErPtNhGLGhV9e7Adehfp";
            string response_type = "token";
            string redirect_uri = "https://www.goodfelow.com/oauth2/callback";
            */
            Scope[] scope = new Scope[]
            { Scope.DataRead, Scope.DataWrite, Scope.DataCreate, Scope.BucketRead, Scope.BucketCreate, Scope.BucketUpdate, Scope.CodeAll, Scope.DataSearch, Scope.ViewablesRead, Scope.BucketDelete};
            string codeurl = oauthree.Authorize(client_id, response_type, redirect_uri, scope);


            //dynamic token = oauthree.Gettoken(client_id, client_secret, oAuthConstants.AUTHORIZATION_CODE, code, redirect_uri);
            //Console.WriteLine(token.access_token);
            //browser = new ChromiumWebBrowser("http://forge.sinotech.com.tw/");
            browser = new ChromiumWebBrowser(codeurl);
            browser.LoadError += Browser_error;

            panel1.Controls.Add(browser);
            browser.Dock = DockStyle.Fill;



        }
        string token_url = "";
        private void Browser_error(object sender, LoadErrorEventArgs e)
        {
            if(e.FailedUrl.Contains("token"))
            {
                token_url = e.FailedUrl.Split('=')[1].Split('&')[0];
                MethodInvoker mi = new MethodInvoker(this.UpdateUI);
                this.BeginInvoke(mi, null);

            }
            else
            {
                MethodInvoker mi = new MethodInvoker(this.UpdateUIF);
                this.BeginInvoke(mi, null);
            }
        }
        private void UpdateUI()
        {
            label1.Text = "登入成功，請點選「上傳」按鈕。";
            panel1.Controls.Clear();
            panel1.Controls.Add(label1);
        }
        private void UpdateUIF()
        {
            label1.Text = "登入失敗，未成功獲取Token，請聯繫維護人員。";
            panel1.Controls.Clear();
            panel1.Controls.Add(label1);
        }



        string s;
        private void button3_Click(object sender, EventArgs e)
        {
            bool b = false;
            try { b = file_tuples[0].Item1 != ""; } catch { }
            if(b)
            {
                this.SendRequest(token_url);

            }
            else
            {
                MessageBox.Show("尚未讀取上傳資訊。");
            }

        }


        public IList<Tuple<string, string, string>> file_tuples = new List<Tuple<string, string, string>>();

        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            //excel
            ExReader dan = new ExReader();
            dan.SetData(openFileDialog.FileName, 4);
            dan.PassFI();
            dan.CloseEx();

            file_tuples = dan.file_information;

        }
        private async void  SendRequest(string tk)
        {


            HubsApi hubapi = new HubsApi();
            hubapi.Configuration.AccessToken = tk;
            //var hubs = hubapi.GetHubs();
            var hubs = await hubapi.GetHubsAsync();
            //MessageBox.Show(hubs.ToString());

            string hubid = "b.bf3f05be-408a-4abd-aff7-f2b765b9a4e8";//中興
            //Console.WriteLine("hub id is " + hubid);

            ProjectsApi proapi = new ProjectsApi();
            proapi.Configuration.AccessToken = tk;
            var projects = proapi.GetHubProjects(hubid) ;
            string projectid = null;
            foreach (KeyValuePair<string, dynamic> projectInfo in new DynamicDictionaryItems(projects.data))
            {
                string aa = projectInfo.Value.attributes.name;

                if (aa == file_tuples[0].Item3)
                {
                    projectid = projectInfo.Value.id.ToString();
                }
            }

            if(projectid == null)
            {
                projectid = "b.709f1949-aa09-4174-8963-4187b4172e33"; //SinoExcavation
            }


            var topfolders = proapi.GetProjectTopFolders(hubid, projectid);
            string topfolderid = null;
            foreach (KeyValuePair<string, dynamic> topfolder in new DynamicDictionaryItems(topfolders.data))
            {
                string topfolder_name = topfolder.Value.attributes.name;

                if(new DynamicDictionaryItems(topfolders.data).Count() == 1)
                {
                    topfolderid = topfolder.Value.id.ToString();
                }
                else if (topfolder_name == "Project Files")
                {
                    topfolderid = topfolder.Value.id.ToString();
                }
            }

            if(topfolderid == null)
            {
                topfolderid = "urn:adsk.wipprod:fs.folder:co.P8Y23I4QRB2Z21_T7VfbEA";//00-ADMIN
            }


            FoldersApi foldersApi = new FoldersApi();
            foldersApi.Configuration.AccessToken = tk;
            var folders = foldersApi.GetFolderContents(projectid, topfolderid);

            string path = file_tuples[0].Item1;

            var folders_path = path.Split('\\');

            string folder_id = "";

            foreach (string name in folders_path)
            {
                if(name != "")
                {

                    folder_id = Search_folder(folders, name);

                    if (folder_id != "null")
                    {
                        folders = foldersApi.GetFolderContents(projectid, folder_id);
                    }
                }
            }

            string filepath = doc_path;

            string filename = file_tuples[0].Item2 + ".rvt";


            // folderid name
            // objectid : included relationships storage data id
            // name : included attributes name

            string storageJson = @"{

                'jsonapi': {'version':'1.0'},
                'data': {
                    'type': 'objects',
                    'attributes': {'name': ''},
                    'relationships': {
                        'target':{
                            'data':{
                                'type': 'folders',
                                'id': ''
                             }
                        }
                    }
                }
            }";
            //data attributes name : filename
            CreateStorage createStorage = JsonConvert.DeserializeObject<CreateStorage>(storageJson);
            createStorage.Data.Attributes.Name = filename;
            createStorage.Data.Relationships.Target.Data.Id = folder_id;

            dynamic postStorage = proapi.PostStorage(projectid, createStorage);
            string bucketkey = "wip.dm.prod";

            string objectID = postStorage["data"]["id"];
            int i = objectID.IndexOf("/");
            string objectName = objectID.Substring(i + 1);

            ObjectsApi objectsApi = new ObjectsApi();
            objectsApi.Configuration.AccessToken = tk;


            dynamic uploadObject;
            FileStream logFileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using (StreamReader sr = new StreamReader(logFileStream))
            {
                uploadObject = objectsApi.UploadObject(bucketkey, objectName, (int)sr.BaseStream.Length, sr.BaseStream);
            }

            try
            {

                ItemsApi itemsApi = new ItemsApi();
                itemsApi.Configuration.AccessToken = tk;
                //upload item
                string itemJson = @"{
                        'jsonapi': { 'version': '1.0' },
                        'data': {
                            'type': 'items',
                            'attributes': {
                                'displayName': 'Have you ever seen the rain',
                                'extension': {
                                    'type': 'items:autodesk.bim360:File',
                                    'version': '1.0'
                                }
                            },
                            'relationships': {
                                'tip': {
                                    'data': {
                                        'type': 'versions', 'id': '1'
                                    }
                                },
                                'parent': {
                                    'data': {
                                        'type': 'folders',
                                        'id': ''
                                    }
                                }
                            }
                        },
                        'included': [
                        {
                            'type': 'versions',
                            'id': '1',
                            'attributes': {
                                'name': '',
                                'extension': {
                                    'type': 'versions:autodesk.bim360:File',
                                    'version': '1.0'
                                }
                            },
                            'relationships': {
                                'storage': {
                                    'data': {
                                        'type': 'objects',
                                        'id': ''
                                    }
                                }
                            }
                        }
                    ]
                }";
                CreateItem createItem = JsonConvert.DeserializeObject<CreateItem>(itemJson);
                createItem.Included[0].Relationships.Storage.Data.Id = objectID;
                createItem.Included[0].Attributes.Name = filename;
                createItem.Data.Relationships.Parent.Data.Id = folder_id;
                dynamic postItem = itemsApi.PostItem(projectid, createItem);
            }
            catch(Exception e)
            {
                //overwrite version

                VersionsApi versionApi = new VersionsApi();
                versionApi.Configuration.AccessToken = tk;
                string itemsid = Search_item(folders, filename);
                string versionJson = @"{
                        'jsonapi': { 'version': '1.0' },
                        'data': {
                            'type': 'versions',
                            'attributes': {
                                'name': '',
                                'extension': { 'type': 'versions:autodesk.bim360:File', 'version': '1.0'}
                                },
                            'relationships': {
                                'item': { 'data': { 'type': 'items', 'id': '' } },
                                'storage': { 'data': { 'type': 'objects', 'id': '' } }
                                    }
                            }
                        }";

                CreateVersion createVersion = JsonConvert.DeserializeObject<CreateVersion>(versionJson);
                createVersion.Data.Attributes.Name = filename;
                createVersion.Data.Relationships.Item.Data.Id = itemsid;
                createVersion.Data.Relationships.Storage.Data.Id = objectID;
                dynamic postVersion = versionApi.PostVersion(projectid, createVersion);
            }
            MessageBox.Show("上傳完畢");

        }


        private string Search_folder(dynamic select_folders, string wanted_name)
        {
            string id = "null";
            //查看所有資料夾
            foreach (KeyValuePair<string, dynamic> folderInfo in new DynamicDictionaryItems(select_folders.data))
            {
                if (folderInfo.Value.attributes.name == wanted_name)
                {
                    id = folderInfo.Value.id;
                }
            }
            return id;
        }
        private string Search_item(dynamic select_folders, string wanted_name)
        {
            string id = "null";
            //查看檔案
            foreach (KeyValuePair<string, dynamic> folderInfo in new DynamicDictionaryItems(select_folders.data))
            {
                if (folderInfo.Value.attributes.displayName == wanted_name)
                {
                    id = folderInfo.Value.id;
                }
            }
            return id;
        }
    }

}
