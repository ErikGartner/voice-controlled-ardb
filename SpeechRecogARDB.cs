namespace Microsoft.Samples.Speech.Recognition.ListBox
{
    using System;
    using System.Drawing;
    using System.Collections;
    using System.ComponentModel;
    using System.Windows.Forms;
    using System.Data;
    using System.Diagnostics;
    using SpeechLib;

     /// <summary>
    ///     This form is a simple test application for the user control 
    ///     defined in the SpeechListBox project.
    /// </summary>
    public class SpeechRecogARDB : System.Windows.Forms.Form
    {
        private System.Windows.Forms.CheckBox chkSpeechEnabled;
        private System.ComponentModel.IContainer components = null;

        private const int grammarId = 10;
        private bool speechEnabled;
        private bool speechInitialized;

        private String PreCommandString = "";
        private SpeechLib.SpInProcRecoContext objRecoContext;
        private SpeechLib.ISpeechRecoGrammar grammar;
        private SpeechLib.ISpeechGrammarRule ruleTopLevel;
        private SpeechLib.ISpeechGrammarRule ruleListItems;
        private SpeechLib.ISpeechGrammarRule ruleNumbers;
        private SpeechLib.ISpeechGrammarRule ruleCommand;

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtNumber;
        private Button button1;
        private Button button2;
        private Button button3;
        private ComboBox comboBox2;
        private Label label5;
        private Button button4;
        private OpenFileDialog openFileDialog1;
        private SaveFileDialog saveFileDialog1;

        public CardLibrary[] libcards;
        public CardLibrary[] cryptcards;
        private ComboBox comboType;
        private ListView listView1;
        private GroupBox groupBox1;
        private Label label1;
        private const String InventoryHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE inventory SYSTEM \"AnarchRevoltInv.dtd\">\n<?xml-stylesheet type=\"text/xsl\" href=\"xsl/inv2html.xsl\"?>\n<inventory formatVersion=\"-TODO-1.0\" databaseVersion=\"-TODO-20040101\" generator=\"Anarch Revolt Deck Builder\">\n";

        public SpeechRecogARDB()
        {

            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            label1.Text = "How To: Enabled voice recognition, use the following syntax:\n<add/set/remove><0-100><[card name]>\nThen press \"Go\" or enter to perform the command, if it was accurate.";
            LoadItems();
            ClearInventory();     
        }

        private void ClearInventory()
        {
          for (int x = 0; x < libcards.Length; x++)
           {
                libcards[x].intCount = 0;
           }

          for (int x = 0; x < cryptcards.Length; x++)
          {
              cryptcards[x].intCount = 0;
          }
          PrepareList();
        }

        private Boolean WriteInventory(String szFile)
        {
            try
            {

                String szBuffer = InventoryHeader;
                szBuffer += "  <date>"+DateTime.Now+"</date>\n";
                szBuffer += "  <crypt size=\"0\">";
                for (int x = 0; x < cryptcards.Length; x++)
                {
                    if (cryptcards[x].intCount > 0)
                    {
                        String card = "";
                        card += "<vampire databaseID=\"" + cryptcards[x].szID + "\" have=\"" + cryptcards[x].intCount + "\" spare=\"0\" need=\"0\">";
                        card += "<name>" + cryptcards[x].szName + "</name>";
                        card += "<set>" + cryptcards[x].szSet + "</set>";
                        card += "<rarity>" + cryptcards[x].szRarity + "</rarity>";
                        card += "<adv>" + cryptcards[x].szAdv + "</adv></vampire>";
                        szBuffer += card;
                    }
                }
                szBuffer += "</crypt>\n";
                
                szBuffer += "  <library size=\"0\">";     //Is 0 for some reason.
                for (int x = 0; x < libcards.Length; x++)
                {
                    if (libcards[x].intCount > 0)
                    {
                        String card = "";
                        card += "<card databaseID=\"" + libcards[x].szID + "\" have=\"" + libcards[x].intCount + "\" spare=\"0\" need=\"0\">";
                        card += "<name>" + libcards[x].szName + "</name>";
                        card += "<set>" + libcards[x].szSet + "</set>";
                        card += "<rarity>" + libcards[x].szRarity + "</rarity></card>";
                        szBuffer += card;
                    }
                }

                szBuffer += "</library>\n</inventory>\n";

                System.IO.File.WriteAllText(szFile, szBuffer); //,Encoding.utf8);
                PrepareList();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error while saving inventory\n\n" + e.ToString());
                PrepareList();
                return false;
            }

        }
        

        //Warning doesn't care about sets, wants and spare fields!
        private Boolean LoadInventory(String szFile)
        {
            
            try
            {
                ClearInventory();
                //Load Inv
                string s = System.IO.File.ReadAllText(szFile);

                //If empty file
                if (s.Length == 0)
                {
                    ClearInventory();
                    return true;
                }

                string[] stringSeparators = new string[] { "<library" };
                //split[0] = crypt + header, split[1] = library
                String[] split = s.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
                
                stringSeparators = new string[] { "<vampire " };
                String[] cards = split[0].Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

                for (int x = 1; x < cards.Length; x++)
                {
                    int start = cards[x].IndexOf("databaseID=\"") + "databaseID=\"".Length;
                    int end = cards[x].IndexOf("\"", start + 1) - start;
                    String szID = cards[x].Substring(start, end);

                    start = cards[x].IndexOf("have=\"") + "have=\"".Length;
                    end = cards[x].IndexOf("\"", start + 1) - start;
                    String szHave = cards[x].Substring(start, end);

                    for (int y = 0; y < cryptcards.Length; y++)
                    {
                        if (cryptcards[y].szID.Equals(szID))
                        {
                            cryptcards[y].intCount = Convert.ToInt32(szHave);
                            break;
                        }
                    }

                }

                //cards
                stringSeparators = new string[] { "<card " };
                cards = split[1].Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);

                for (int x = 1; x < cards.Length; x++)
                {

                    int start = cards[x].IndexOf("databaseID=\"") + "databaseID=\"".Length;
                    int end = cards[x].IndexOf("\"", start + 1) - start;
                    String szID = cards[x].Substring(start, end);

                    start = cards[x].IndexOf("have=\"") + "have=\"".Length;
                    end = cards[x].IndexOf("\"", start + 1) - start;
                    String szHave = cards[x].Substring(start, end);

                    for (int y = 0; y < libcards.Length; y++)
                    {
                        if (libcards[y].szID.Equals(szID))
                        {
                            libcards[y].intCount = Convert.ToInt32(szHave);
                            break;
                        }
                    }

                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error while loading inventory\n\n" + e.ToString());
                PrepareList();
                return false;
            }
        }


        //Reads the cards to the combobox
        public void PrepareList()
        {

            listView1.Clear();
            switch (comboType.SelectedIndex)
            {
                case -1:
                    comboType.SelectedIndex = 1;
                    PrepareList();
                    break;

                case 0:
                    listView1.Columns.Add("Have");
                    listView1.Columns.Add("Vampire");
                    listView1.Columns.Add("Adv");
                    for (int x = 0; x < cryptcards.Length; x++)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Text = cryptcards[x].intCount.ToString();                        
                        lvi.SubItems.Add(cryptcards[x].szName);
                        lvi.SubItems.Add(cryptcards[x].szAdv);
                        listView1.Items.Add(lvi);
                    }
                    break;

                case 1:
                    listView1.Columns.Add("Have");
                    listView1.Columns.Add("Card");
                    for (int x = 0; x < libcards.Length; x++)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Text = libcards[x].intCount.ToString();
                        lvi.SubItems.Add(libcards[x].szName);
                        listView1.Items.Add(lvi);
                    }
                    break;
            }        

        }
        
        public void LoadItems()
        {
            String s = System.IO.File.ReadAllText("lib.db");
            s = s.Replace("\r","");
            char[] chrseperator = new char[] { '\n' };
            string[] items = s.Split(chrseperator, StringSplitOptions.RemoveEmptyEntries);

            string[] stringSeparators = new string[] { "###" };
            libcards = new CardLibrary[items.Length];

            for (int x = 0; x < items.Length; x++)
            {
                //NAME###ID###SET###RARITY
                String[] values = items[x].Split(stringSeparators, StringSplitOptions.None);
                CardLibrary card = new CardLibrary();

                card.szName = values[0];
                card.szID = values[1];
                card.szSet = values[2];
                card.szRarity = values[3];

                libcards[x] = card;

            }

            s = System.IO.File.ReadAllText("crypt.db");
            s = s.Replace("\r", "");
            chrseperator = new char[] { '\n' };
            items = s.Split(chrseperator, StringSplitOptions.RemoveEmptyEntries);

            stringSeparators = new string[] { "###" };
            cryptcards = new CardLibrary[items.Length];

            for (int x = 0; x < items.Length; x++)
            {
                //NAME###ID###SET###RARITY###ADV
                String[] values = items[x].Split(stringSeparators, StringSplitOptions.None);
                CardLibrary card = new CardLibrary();

                card.szName = values[0];
                card.szID = values[1];
                card.szSet = values[2];
                card.szRarity = values[3];
                card.szAdv = values[4];

                cryptcards[x] = card;

            }         

        }

        /*Speech recog methods*/
        /// <summary>
        ///     RecoContext_Hypothesis is the event handler function for 
        ///     SpInProcRecoContext object's Hypothesis event.
        /// </summary>
        /// <param name="StreamNumber"></param>
        /// <param name="StreamPosition"></param>
        /// <param name="Result"></param>
        /// <remarks>
        ///     See EnableSpeech() for how to hook up this function with the 
        ///     event.
        /// </remarks>
        private void RecoContext_Hypothesis(int StreamNumber,
            object StreamPosition,
            ISpeechRecoResult Result)
        {
            Debug.WriteLine("Hypothesis: " +
                Result.PhraseInfo.GetText(0, -1, true) + ", " +
                StreamNumber + ", " + StreamPosition);
        }

        /// <summary>
        ///     RecoContext_Hypothesis is the event handler function for 
        ///     SpInProcRecoContext object's Recognition event.
        /// </summary>
        /// <param name="StreamNumber"></param>
        /// <param name="StreamPosition"></param>
        /// <param name="RecognitionType"></param>
        /// <param name="Result"></param>
        /// <remarks>
        ///     See EnableSpeech() for how to hook up this function with the 
        ///     event.
        /// </remarks>
        private void RecoContext_Recognition(int StreamNumber, object StreamPosition, SpeechRecognitionType RecognitionType, ISpeechRecoResult Result)
        {
            Debug.WriteLine("Recognition: " + Result.PhraseInfo.GetText(0, -1, true) + ", " + StreamNumber + ", " + StreamPosition);

            int index;
            int index1;
            int index2;
            ISpeechPhraseProperty oCard;
            ISpeechPhraseProperty oNumber;
            ISpeechPhraseProperty oCommand;

            // oItem will be the property of the second part in the recognized 
            // phase. For example, if the top level rule matchs 
            // "select Seattle". Then the ListItemsRule matches "Seattle" part.
            // The following code will get the property of the "Seattle" 
            // phrase, which is set when the word "Seattle" is added to the 
            // ruleListItems in RebuildGrammar.
            oCommand = Result.PhraseInfo.Properties.Item(0).Children.Item(0);
            index = oCommand.Id;

            oNumber = Result.PhraseInfo.Properties.Item(1).Children.Item(0);
            index1 = oNumber.Id;

            oCard = Result.PhraseInfo.Properties.Item(2).Children.Item(0);
            index2 = oCard.Id;

            if ((System.Decimal)Result.PhraseInfo.GrammarId == grammarId)
            {
                // Check to see if the item at the same position in the list 
                // still has the same text.
                // This is to prevent the rare case that the user keeps 
                // talking while the list is being added or removed. By the 
                // time this event is fired and handled, the list box may have 
                // already changed.
                if (oCard.Name.CompareTo(libcards[index2].ToString()) == 0||oCard.Name.CompareTo(cryptcards[index2].ToString())==0)
                {
                    listView1.Items[index2].Selected = true;
                    listView1.Items[index2].Focused = true;
                    listView1.TopItem = listView1.Items[index2];                    
                    txtNumber.Text = oNumber.Name;
                    comboBox2.SelectedIndex = index;
                }
            }
        }

        /// <summary>
        ///     This function will create the main SpInProcRecoContext object 
        ///     and other required objects like Grammar and rules. 
        ///     In this sample, we are building grammar dynamically since 
        ///     listbox content can change from time to time.
        ///     If your grammar is static, you can write your grammar file 
        ///     and ask SAPI to load it during run time. This can reduce the 
        ///     complexity of your code.
        /// </summary>
        private void InitializeSpeech()
        {
            Debug.WriteLine("Initializing SAPI objects...");

            try
            {
                // First of all, let's create the main reco context object. 
                // In this sample, we are using inproc reco context. Shared reco
                // context is also available. Please see the document to decide
                // which is best for your application.
                objRecoContext = new SpeechLib.SpInProcRecoContext();

                SpeechLib.SpObjectTokenCategory objAudioTokenCategory = new SpeechLib.SpObjectTokenCategory();
                objAudioTokenCategory.SetId(SpeechLib.SpeechStringConstants.SpeechCategoryAudioIn, false);

                SpeechLib.SpObjectToken objAudioToken = new SpeechLib.SpObjectToken();
                objAudioToken.SetId(objAudioTokenCategory.Default, SpeechLib.SpeechStringConstants.SpeechCategoryAudioIn, false);

                objRecoContext.Recognizer.AudioInput = objAudioToken;

                // Then, let's set up the event handler. We only care about
                // Hypothesis and Recognition events in this sample.
                //objRecoContext.Hypothesis += new _ISpeechRecoContextEvents_HypothesisEventHandler(RecoContext_Hypothesis);

                objRecoContext.Recognition += new _ISpeechRecoContextEvents_RecognitionEventHandler(RecoContext_Recognition);

                // Now let's build the grammar.
                // The top level rule consists of two parts: "select <items>". 
                // So we first add a word transition for the "select" part, then 
                // a rule transition for the "<items>" part, which is dynamically 
                // built as items are added or removed from the listbox.
                grammar = objRecoContext.CreateGrammar(grammarId);
                ruleTopLevel = grammar.Rules.Add("TopLevelRule", SpeechRuleAttributes.SRATopLevel | SpeechRuleAttributes.SRADynamic, 1);
                ruleCommand = grammar.Rules.Add("CommandRule", SpeechRuleAttributes.SRADynamic, 2);
                ruleNumbers = grammar.Rules.Add("NumberRule", SpeechRuleAttributes.SRADynamic, 3);
                ruleListItems = grammar.Rules.Add("ListItemsRule", SpeechRuleAttributes.SRADynamic, 4);
                
                //Prepare states
                SpeechLib.ISpeechGrammarRuleState stateAfterPre;
                SpeechLib.ISpeechGrammarRuleState stateAfterCommand;                
                SpeechLib.ISpeechGrammarRuleState stateAfterNumber;
                stateAfterPre = ruleTopLevel.AddState();
                stateAfterCommand = ruleTopLevel.AddState();
                stateAfterNumber = ruleTopLevel.AddState();
                
                //Add keywords: add,set,delete
                object PropValue = "";
                ruleTopLevel.InitialState.AddWordTransition(stateAfterPre, PreCommandString, null, SpeechGrammarWordType.SGLexicalNoSpecialChars, "", 0, ref PropValue, 1.0F);

                String word;
                PropValue = "";
                word = "Add";
                ruleCommand.InitialState.AddWordTransition(null, word, "", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, 0, ref PropValue, 1f);

                word = "Set";
                ruleCommand.InitialState.AddWordTransition(null, word, "", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, 1, ref PropValue, 1f);

                word = "Remove";
                ruleCommand.InitialState.AddWordTransition(null, word, "", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, 2, ref PropValue, 1f);

                stateAfterPre.AddRuleTransition(stateAfterCommand, ruleCommand, "", 1, ref PropValue, 1F);

                PropValue = "";
                for (int x = 0; x <= 100; x++)
                {

                    word = Convert.ToString(x);

                    // Note: if the same word is added more than once to the same 
                    // rule state, SAPI will return error. In this sample, we 
                    // don't allow identical items in the list box so no need for 
                    // the checking, otherwise special checking for identical words
                    // would have to be done here.
                    ruleNumbers.InitialState.AddWordTransition(null, word, "", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, x, ref PropValue, 1F);
                }

                stateAfterCommand.AddRuleTransition(stateAfterNumber, ruleNumbers, "", 2, ref PropValue, 1.0F);

                PropValue = "";
                stateAfterNumber.AddRuleTransition(null, ruleListItems, "", 3, ref PropValue, 1.0F);

                // Now add existing list items to the ruleListItems
                RebuildGrammar();

                // Now we can activate the top level rule. In this sample, only 
                // the top level rule needs to activated. The ListItemsRule is 
                // referenced by the top level rule.
                grammar.CmdSetRuleState("TopLevelRule", SpeechRuleState.SGDSActive);
                speechInitialized = true;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Exception caught when initializing SAPI."
                    + " This application may not run correctly.\r\n\r\n"
                    + e.ToString(),
                    "Error");

                throw;
            }
        }

        /// <summary>
        ///     EnableSpeech will initialize all speech objects on first time,
        ///     then rebuild grammar and start speech recognition.
        /// </summary>
        /// <returns>
        ///     true if speech is enabled and grammar updated.
        ///     false otherwise, which happens if we are in design mode.
        /// </returns>
        /// <remarks>
        ///     This is a private function.
        /// </remarks>
        private bool EnableSpeech()
        {
            Debug.Assert(speechEnabled, "speechEnabled must be true in EnableSpeech");

            if (this.DesignMode) return false;

            if (speechInitialized == false)
            {
                this.Enabled = false;
                InitializeSpeech();
                this.Enabled = true;
            }
            else
            {
                this.Enabled = false;
                RebuildGrammar();
                this.Enabled = true;
            }

            objRecoContext.State = SpeechRecoContextState.SRCS_Enabled;
            return true;
        }

        /// <summary>
        ///     RebuildGrammar() will update grammar object with current list 
        ///     items. It is called automatically by AddItem and RemoveItem.
        /// </summary>
        /// <returns>
        ///     true if grammar is updated.
        ///     false if grammar is not updated, which can happen if speech is 
        ///     not enabled or if it's in design mode.
        /// </returns>
        /// <remarks>
        ///     RebuildGrammar should be called every time after the list item 
        ///     has changed. AddItem and RemoveItem methods are provided as a 
        ///     way to update list item and the grammar object automatically.
        ///     Don't forget to call RebuildGrammar if the list is changed 
        ///     through ListBox.Items collection. Otherwise speech engine will 
        ///     continue to recognize old list items.
        /// </remarks>
        public bool RebuildGrammar()
        {

            if (comboType.SelectedIndex == -1)
            {
                return false;
            }

            if (!speechEnabled || this.DesignMode)
            {
                return false;
            }

            // In this funtion, we are only rebuilding the ruleListItems, as 
            // this is the only part that's really changing dynamically in 
            // this sample. However, you still have to call 
            // Grammar.Rules.Commit to commit the grammar.
            int i;
            String word;
            object propValue = "";

            try
            {

                // Note: if the same word is added more than once to the same 
                // rule state, SAPI will return error. In this sample, we 
                // don't allow identical items in the list box so no need for 
                // the checking, otherwise special checking for identical words
                // would have to be done here.
                ruleListItems.Clear();
                switch (comboType.SelectedIndex)
                {
                    case 0:
                        for (i = 0; i < cryptcards.Length; i++)
                        {
                            word = cryptcards[i].ToString();
                            ruleListItems.InitialState.AddWordTransition(null, word, " ", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, i, ref propValue, 1F);
                        }
                        break;
                    case 1:
                        for (i = 0; i < libcards.Length; i++)
                        {
                            word = libcards[i].ToString();
                            ruleListItems.InitialState.AddWordTransition(null, word, " ", SpeechGrammarWordType.SGLexicalNoSpecialChars, word, i, ref propValue, 1F);
                        }
                        break;
                }
                grammar.Rules.Commit();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Exception caught when rebuilding dynamic ListBox rule.\r\n\r\n"
                    + e.ToString(),
                    "Error");

                throw;
            }

            return true;
        }


        /// <summary>
        ///     This is a private function that stops speech recognition.
        /// </summary>
        /// <returns></returns>
        private bool DisableSpeech()
        {
            if (this.DesignMode) return false;

            Debug.Assert(speechInitialized,
                         "speech must be initialized in DisableSpeech");

            if (speechInitialized)
            {
                // Putting the recognition context to disabled state will 
                // stop speech recognition. Changing the state to enabled 
                // will start recognition again.
                objRecoContext.State = SpeechRecoContextState.SRCS_Disabled;
            }

            return true;
        }

        /// <summary>
        ///     Property SpeechEnabled is read/write-able. When it's set to
        ///     true, speech recognition will be started. When it's set to
        ///     false, speech recognition will be stopped.
        /// </summary>
        public bool SpeechEnabled
        {
            get
            {
                return speechEnabled;
            }
            set
            {
                if (speechEnabled != value)
                {
                    speechEnabled = value;
                    if (this.DesignMode) return;

                    if (speechEnabled)
                    {
                        EnableSpeech();
                    }
                    else
                    {
                        DisableSpeech();
                    }
                }
            }
        }

        /// <summary>
        ///     Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.chkSpeechEnabled = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtNumber = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.comboType = new System.Windows.Forms.ComboBox();
            this.listView1 = new System.Windows.Forms.ListView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkSpeechEnabled
            // 
            this.chkSpeechEnabled.Location = new System.Drawing.Point(267, 289);
            this.chkSpeechEnabled.Name = "chkSpeechEnabled";
            this.chkSpeechEnabled.Size = new System.Drawing.Size(112, 16);
            this.chkSpeechEnabled.TabIndex = 2;
            this.chkSpeechEnabled.Text = "&Speech Enabled";
            this.chkSpeechEnabled.CheckedChanged += new System.EventHandler(this.chkSpeechEnabled_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(85, 271);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Number";
            // 
            // txtNumber
            // 
            this.txtNumber.Location = new System.Drawing.Point(88, 287);
            this.txtNumber.Name = "txtNumber";
            this.txtNumber.Size = new System.Drawing.Size(64, 20);
            this.txtNumber.TabIndex = 6;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(158, 287);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(103, 22);
            this.button1.TabIndex = 8;
            this.button1.Text = "Go";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(129, 19);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(103, 23);
            this.button2.TabIndex = 9;
            this.button2.Text = "Save";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(252, 19);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(103, 23);
            this.button3.TabIndex = 10;
            this.button3.Text = "Load";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "Add",
            "Set",
            "Remove"});
            this.comboBox2.Location = new System.Drawing.Point(10, 287);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(72, 21);
            this.comboBox2.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 271);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(54, 13);
            this.label5.TabIndex = 14;
            this.label5.Text = "Command";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(6, 19);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(103, 23);
            this.button4.TabIndex = 15;
            this.button4.Text = "New";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // comboType
            // 
            this.comboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboType.FormattingEnabled = true;
            this.comboType.Items.AddRange(new object[] {
            "Crypt",
            "Library"});
            this.comboType.Location = new System.Drawing.Point(10, 12);
            this.comboType.Name = "comboType";
            this.comboType.Size = new System.Drawing.Size(361, 21);
            this.comboType.TabIndex = 16;
            this.comboType.SelectedIndexChanged += new System.EventHandler(this.comboType_SelectedIndexChanged);
            // 
            // listView1
            // 
            this.listView1.FullRowSelect = true;
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(10, 39);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(361, 227);
            this.listView1.TabIndex = 17;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button4);
            this.groupBox1.Location = new System.Drawing.Point(10, 315);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(361, 61);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Inventory";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 389);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 19;
            this.label1.Text = "Placeholder";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // SpeechRecogARDB
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(383, 438);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.comboType);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtNumber);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkSpeechEnabled);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(399, 476);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(399, 476);
            this.Name = "SpeechRecogARDB";
            this.ShowIcon = false;
            this.Text = "POC: Voice Controlled ARDB Inventory, by SmoiZ";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() 
        {
            Application.Run(new SpeechRecogARDB());
        }
        
      
        private void chkSpeechEnabled_CheckedChanged(object sender, System.EventArgs e)
        {
            if (chkSpeechEnabled.Checked)
            {
                SpeechEnabled = true; ;
            }
            else
            {
                SpeechEnabled = false; ;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int no = Convert.ToInt32(txtNumber.Text);
                if (no < 1)
                {
                    return;
                }

                if (listView1.SelectedItems.Count == 0)
                {
                    return;
                }

                if (comboBox2.SelectedItem == null)
                {
                    return;
                }

                CardLibrary card;
                switch(comboType.SelectedIndex){
                    case 0:
                        card = (CardLibrary)cryptcards[listView1.SelectedItems[0].Index];
                        break;

                    case 1:
                        card = (CardLibrary)libcards[listView1.SelectedItems[0].Index];
                        break;

                    default:
                        return;
                }

                switch (comboBox2.SelectedIndex)
                {

                        //add
                    case 0:
                        card.intCount += no;
                        break;

                        //set
                    case 1:
                        card.intCount = no;
                        break;

                        //remove
                    case 2:
                        card.intCount -= no;
                        if (card.intCount < 0)
                        {
                            card.intCount = 0;
                        }
                        break;

                }
                listView1.SelectedItems[0].SubItems[0].Text = card.intCount.ToString();
                txtNumber.Text = "";
                listView1.SelectedItems.Clear();
            }
            catch (Exception excp)
            {
                return;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(saveFileDialog1.ShowDialog()!= DialogResult.OK){
                return;
            }
            WriteInventory(saveFileDialog1.FileName);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            LoadInventory(openFileDialog1.FileName);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ClearInventory();
        }

        private void comboType_SelectedIndexChanged(object sender, EventArgs e)
        {
            PrepareList();
            RebuildGrammar();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

       
    }

    public class CardLibrary
    {
        public String szName;
        public String szID;
        public String szRarity;
        public String szSet;
        public String szAdv; //Advanced - only crypt
        public int intCount = 0;    //antal

        public override string ToString()
        {
            return szName;
        }

    }
}
