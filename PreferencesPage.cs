using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using BetterReminders.Properties;
using Outlook = Microsoft.Office.Interop.Outlook;

// Copyright (c) 2016-2017, 2019 Ben Spiller.

namespace BetterReminders
{
    [ComVisible(true)]
    public partial class PreferencesPage : UserControl, Outlook.PropertyPage
    {
        private Logger logger = Logger.GetLogger();
        private Outlook.PropertyPageSite propertyPageSite;

        public PreferencesPage()
        {
            InitializeComponent();
        }

        bool isDirty;
        void Outlook.PropertyPage.Apply()
        {
            string meetingregex = meetingUrlRegex.Text;
            // normalize use of default regex to improve upgradeability
            if (meetingregex == UpcomingMeeting.DefaultMeetingUrlRegex)
                meetingregex = "";
            if (meetingregex != "")
                try
                {
                    Regex re = new Regex(meetingregex);
                    if (!re.GetGroupNames().Contains("url"))
                        throw new Exception($"The meeting regex must include a regex group named 'url' e.g. '{UpcomingMeeting.DefaultMeetingUrlRegex}'");
                } catch (Exception e)
                {
                    string msg = $"Invalid meeting URL regex: {e.Message}";
                    MessageBox.Show(msg, "Invalid Meeting URL Regex", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw new Exception(msg, e); // stops isDirty being changed
                }

            string subjectexcluderegex = subjectExcludeRegex.Text;
            if (subjectexcluderegex != "")
                try
                {
                    Regex re = new Regex(subjectexcluderegex);
                }
                catch (Exception e)
                {
                    string msg = $"Invalid subject exclude regex: {e.Message}";
                    MessageBox.Show(msg, "Invalid Subject Exclude Regex", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw new Exception(msg, e); // stops isDirty being changed
                }

            // first, validation
            string reminderSound = (reminderSoundPath.Text == "(none)") ? "" : reminderSoundPath.Text;
            if (reminderSound != "" && reminderSound != "(default)" &&
                !System.IO.File.Exists(reminderSound))
            {
                MessageBox.Show("Reminder .wav path does not exist. Provide a valid .wav path, empty string or (default).",
                                "Invalid BetterReminders settings",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                throw new Exception("BetterReminders got invalid input"); // stops isDirty being changed
            }

            Settings.Default.defaultReminderSecs = decimal.ToInt32(defaultReminderTimeSecs.Value);
            Settings.Default.searchFrequencySecs = decimal.ToInt32(searchFrequencyMins.Value) * 60;
            Settings.Default.playSoundOnReminder = reminderSound;
            Settings.Default.meetingUrlRegex = meetingregex;
            Settings.Default.subjectExcludeRegex = subjectexcluderegex;
            Settings.Default.Save();
            isDirty = false;
        }

        bool Outlook.PropertyPage.Dirty => isDirty;
        void Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {
            // nothing to do here
        }

        private void valueChanged(object sender, EventArgs e)
        {
            if (propertyPageSite == null || isDirty) return; // still loading or already called

            isDirty = true;
            propertyPageSite.OnStatusChange();
        }

        private Outlook.PropertyPageSite GetPropertyPageSite()
        {
            // this is what MS's documentation recommends, but doesn't seem to work as Parent is null
            if (Parent is Outlook.PropertyPageSite propertyPageSite2) return propertyPageSite2;

            // nb: I can't believe this hack is really required, but since Parent=null
            // I can't find any better way to do it

            Type type = typeof(object);
            string assembly = type.Assembly.CodeBase.Replace("mscorlib.dll", "System.Windows.Forms.dll");
            assembly = assembly.Replace("file:///", "");

            string assemblyName = AssemblyName.GetAssemblyName(assembly).FullName;
            Type unsafeNativeMethods = Type.GetType(Assembly.CreateQualifiedName(assemblyName, "System.Windows.Forms.UnsafeNativeMethods"));

            Type oleObj = unsafeNativeMethods.GetNestedType("IOleObject");
            MethodInfo methodInfo = oleObj.GetMethod("GetClientSite");

            return methodInfo.Invoke(this, null) as Outlook.PropertyPageSite;
        }

        private void PreferencesPage_Load(object sender, EventArgs e)
        {
            try
            {
                defaultReminderTimeSecs.Value = Settings.Default.defaultReminderSecs;
                searchFrequencyMins.Value = Math.Max(1, Math.Min(Settings.Default.searchFrequencySecs/60, searchFrequencyMins.Maximum));
                reminderSoundPath.Text = Settings.Default.playSoundOnReminder;

                // provide default in case user forgets
                meetingUrlRegex.Items.Add(UpcomingMeeting.DefaultMeetingUrlRegex);

                meetingUrlRegex.Text = Settings.Default.meetingUrlRegex;
                if (string.IsNullOrWhiteSpace(meetingUrlRegex.Text))
                    meetingUrlRegex.Text = UpcomingMeeting.DefaultMeetingUrlRegex;

                subjectExcludeRegex.Text = Settings.Default.subjectExcludeRegex;

                propertyPageSite = GetPropertyPageSite();
                logger.Info("Successfully loaded preferences page");
            }
            catch (Exception ex)
            {
                logger.Error("Error loading preferences page: ", ex);
                throw;
            }
        }

        private void reminderSoundBrowse_Click(object sender, EventArgs e)
        {
            reminderSoundBrowseDialog.FileName = reminderSoundPath.Text.StartsWith("(")
                ? ""
                : reminderSoundPath.Text;
            if (reminderSoundBrowseDialog.ShowDialog(ParentForm) == DialogResult.OK)
                reminderSoundPath.Text = reminderSoundBrowseDialog.FileName;
        }
    }
}
