using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using AjaxControlToolkit;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Threading;
//using cisf_Mail;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

using System.Security;

public partial class _Default : System.Web.UI.Page
{
    private static string smtpServer,
                              sendToUser,
                              sendToAdmin,
                              sendFrom,
                              sharepointlibrary,
                              slebbfolder,
                              uploadfolder,
                              jpgfolder,
                              ppsxfolder,
                              spdomain,
                              spname,
                              sppassword,
                              sharedomain,
                              sharename,
                              sharepassword,
                              thispptx,
                              thisppsx,
                              thispptxfullpath,
                              thisppsxfullpath;

    private static FileInfo[] files, jpgfiles;
    private static string[] slebbs;

    public _Default()
    {
        smtpServer = safeGetAppSetting("smtpServer");
        sendToUser = safeGetAppSetting("sendToUser");
        sendToAdmin = safeGetAppSetting("sendToAdmin");
        sendFrom = safeGetAppSetting("sendFrom");
        sharepointlibrary = safeGetAppSetting("SharePointLibrary");
        slebbfolder = safeGetAppSetting("SLEBBFolder");
        uploadfolder = safeGetAppSetting("UploadFolder");
        jpgfolder = safeGetAppSetting("JPGFolder");
        ppsxfolder = safeGetAppSetting("PPSXFolder");
        spdomain = safeGetAppSetting("spDomain");
        spname = safeGetAppSetting("spName");
        sppassword = safeGetAppSetting("spPwd");
        sharedomain = safeGetAppSetting("shareDomain");
        sharename = safeGetAppSetting("shareName");
        sharepassword = safeGetAppSetting("sharePwd"); 
        slebbs = safeGetAppSetting("SLEBBs").Split(new Char[] { ',' });
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            CleanUpFolders();
            lblError.Text = "Click 'Browse', select your PowerPoint presentation (.pptx) then 'Upload'.";
        }
    }
    
    protected void CleanUpFolders()
    {
        CleanUpFolder(Server.MapPath("~/" + uploadfolder));
        CleanUpFolder(Server.MapPath("~/" + jpgfolder));
        CleanUpFolder(Server.MapPath("~/" + ppsxfolder));
    }

    protected void CleanUpFolder(string folder)
    {
        string[] path = Directory.GetFiles(folder);
        if (path != null)
        {
            try
            {
                foreach (string deadfile in path){File.Delete(deadfile);}
            }
            catch (Exception ex6)
            {
                lblError.Text = "Cleanup Error: " + ex6.Message;
            }
        }
    }

    private void ShowDisplays()
    {
        //DirectoryInfo dir = new DirectoryInfo(slebbfolder);
        //slebbs = dir.GetFiles("*.ppsx");
        ArrayList list = new ArrayList();
        list.Add("CISF Portal");
        foreach (string ebb in slebbs)
        {
            list.Add("Rm. " + ebb);
        }
        cbxDirStat.Visible = true;
        cbxDirStat.DataSource = list;
        cbxDirStat.DataBind();

        btnFileProcessing.Visible = true;
        lblError.Visible = true;
    }

    protected void CheckAllBoxes(object sender, EventArgs e)
    {
        foreach (ListItem li in cbxDirStat.Items)
        {
            li.Selected = true;
        }
    }

    protected void UploadPPTX(object sender, EventArgs e)
    {
        if (!fileupload1.PostedFile.FileName.ToLower().EndsWith(".pptx"))
        { lblError.Text = "That file is not a PowerPoint presentation.<br/>I'm looking for any file that ends in .pptx"; }
        else
        {
            // Checks if there is a file present for uploading.
            if ((fileupload1.PostedFile != null) && (fileupload1.PostedFile.ContentLength > 0))
            {
                lblError.Text = string.Empty;
                lblError.Text = string.Empty;
                thispptx = System.IO.Path.GetFileName(fileupload1.PostedFile.FileName);
                thispptxfullpath = Server.MapPath("~/upload/") + thispptx;
                try
                {
                    // Saving to Files folder
                    fileupload1.PostedFile.SaveAs(thispptxfullpath);
                }
                catch (Exception ex)
                {
                    lblError.Text = "File upload error.<br />";
                }
            }
            else
            {
                lblError.Text += "Please select a file to upload.";
            }
            // Deletes hidden sides
            DeleteHiddenSlides(thispptxfullpath);

            // Used for displaying slides to uploaded to sharepoint
            ExtractJPGs(thispptxfullpath);

            // convert to ppsx
            SavePresentationAsSlideshow(thispptxfullpath);

            ShowDisplays();

            // Displays images for review
            ListImages();

            lblError.Text = fileupload1.PostedFile.FileName + " was uploaded and hidden slides were removed.";
            pnlImages.Visible = true;
            pnlSLEBBs.Visible = true;
            pnlUpload.Visible = false;
        }

    }

    public void DeleteHiddenSlides(string pppath)
    {
        int slideCount = -1;

        slideCount = CountSlides(pppath);
        lblError.Text += slideCount + " slides processed.<br />";

        for (int i = 0; i <= slideCount; i++)
        {
            DeleteSlide(pppath);
        }
    }

    public static void DeleteSlide(string presentationFile)
    {
        if (presentationFile == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            // Get the presentation part from the presentation document. 
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Get the presentation from the presentation part.
            DocumentFormat.OpenXml.Presentation.Presentation presentation = presentationPart.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation.SlideIdList;

            int slideIdx = -1;
            foreach (SlideId _slideId in presentation.SlideIdList)
            {
                slideIdx++;

                // Get the slide ID of the specified slide
                SlideId slideId = slideIdList.ChildElements[slideIdx] as SlideId;

                // Get the relationship ID of the slide.
                string slideRelId = slideId.RelationshipId;

                // Get the slide part for the specified slide.
                SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

                if (slidePart.Slide.Show != null)
                {
                    if (slidePart.Slide.Show.HasValue != null)
                    {
                        // Remove the slide from the slide list.
                        slideIdList.RemoveChild(slideId);

                        // Remove references to the slide from all custom shows.
                        if (presentation.CustomShowList != null)
                        {
                            // Iterate through the list of custom shows.
                            foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                            {
                                if (customShow.SlideList != null)
                                {
                                    // Declare a link list of slide list entries.
                                    LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                                    {
                                        // Find the slide reference to remove from the custom show.
                                        if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                                        {
                                            slideListEntries.AddLast(slideListEntry);
                                        }
                                    }

                                    // Remove all references to the slide from the custom show.
                                    foreach (SlideListEntry slideListEntry in slideListEntries)
                                    {
                                        customShow.SlideList.RemoveChild(slideListEntry);
                                    }
                                }
                            }
                        }

                        // Save the modified presentation.
                        presentation.Save();

                        // Remove the slide part.
                        presentationPart.DeletePart(slidePart);
                        break;
                    }
                }
            }

        }
    }

    public static int CountSlides(string presentationFile)
    {
        // Open the presentation as read-only.
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
        {
            // Pass the presentation to the next CountSlide method
            // and return the slide count.
            return CountSlides(presentationDocument);
        }
    }

    public static int CountSlides(PresentationDocument presentationDocument)
    {
        // Check for a null document object.
        if (presentationDocument == null)
        {
            throw new ArgumentNullException("presentationDocument");
        }

        int slidesCount = 0;

        // Get the presentation part of document.
        PresentationPart presentationPart = presentationDocument.PresentationPart;

        // Get the slide count from the SlideParts.
        if (presentationPart != null)
        {
            slidesCount = presentationPart.SlideParts.Count();
        }

        // Return the slide count to the previous method.
        return slidesCount;
    }

    public void SavePresentationAsSlideshow(string presentationFile)
    {
        using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
        {
            try
            {
                presentationDocument.ChangeDocumentType(PresentationDocumentType.Slideshow);
                presentationDocument.Close();
                File.Copy(Server.MapPath("~/upload/default.pptx"), Server.MapPath("~/ppsx/default.ppsx"));
//                File.Copy(Server.MapPath("~/upload")"c:\\_websites\\cisfslebbmanagement\\upload\\default.pptx", @"c:\\_websites\\cisfslebbmanagement\\ppsx\\default.ppsx");
            }
            catch (Exception ey)
            {
                lblError.Text = "Save Presentation As Slideshow Error: " + ey.Message;
            }
        }
    }

    protected void ExtractJPGs(string pppath)
    {
        using (new Impersonate.Impersonation(sharedomain, sharename, sharepassword))
        {
            Microsoft.Office.Interop.PowerPoint.Application ppApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentation ppPresentation = ppApplication.Presentations.Open(Server.MapPath("~/upload/Default.pptx"), MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            try
            {
                ppPresentation.SaveAs(Server.MapPath("~/jpg"), PpSaveAsFileType.ppSaveAsJPG, MsoTriState.msoFalse);
                Thread.Sleep(500);

            }
            catch (Exception ex2)
            {
                lblError.Text = "ExtractJPGs Error: " + ex2.Message;
            }
            finally
            {
                ppPresentation.Close();
            }
        }
    }

    private void ListImages()
    {
        DirectoryInfo dir = new DirectoryInfo(MapPath("~/jpg"));
        files = dir.GetFiles();
        ArrayList list = new ArrayList();
        foreach (FileInfo file in files)
        {
            if (file.Extension == ".JPG")
            {
                list.Add(file);
            }
        }
        dlImages.Visible = true;
        dlImages.DataSource = list;
        dlImages.DataBind();
    }

    protected void UpdateBBs_OnClick(object sender, EventArgs e)
    {
        string msg = "";
        foreach (ListItem li in cbxDirStat.Items)
        {
            if (li.Selected == true)
            {
                if (li.Text == "CISF Portal")
                {
                    msg += CopyToSharePointLibrary();                    
                }else{
                    thisppsx = li.Text.Replace("Rm. ", "Default") + ".ppsx";
                    File.Copy(Server.MapPath("~/" + ppsxfolder) + @"\\default.ppsx", Path.Combine(slebbfolder, thisppsx), true);
                    msg += "<br />" + li.Text + " was updated.";
                }
            }
        }
        lblError.Text = msg;
        CleanUpFolders();
        pnlImages.Visible = false;
        pnlSLEBBs.Visible = false;
        pnlUpload.Visible = true;
    }

    public string CopyToSharePointLibrary()
    {
        string msg = "";
        using (new Impersonate.Impersonation(spdomain, spname, sppassword))
        {
            try
            {
                Array.ForEach(Directory.GetFiles(sharepointlibrary), File.Delete);

                DirectoryInfo dir = new DirectoryInfo(MapPath("~/jpg"));
                jpgfiles = dir.GetFiles();
                foreach (FileInfo jpgfile in jpgfiles)
                {
                    File.Copy(jpgfile.FullName, Path.Combine(sharepointlibrary, jpgfile.Name), true);
                }
                msg = "<br />CISF Portal was updated.";
            }
            catch (Exception ex6_5)
            {
                msg = "<br />Copy To SharePoint Library Error: " + ex6_5.Message;
            }
        }
        return msg;
    }

    protected void Cancel_OnClick(object sender, EventArgs e)
    {
        Response.Redirect("default.aspx");
    }

    private string safeGetAppSetting(string key)
    {
        try
        {
            return ConfigurationManager.AppSettings[key];
        }
        catch (ConfigurationErrorsException ex)
        {
            lblError.Text = "Web.Config Error: " + key + " " + ex.Message;
            return "Key is missing " + key;
        }
    }
}