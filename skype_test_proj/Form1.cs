using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SKYPE4COMLib;
namespace skype_test_proj
{

    public partial class Form1 : Form
    {
        private static Skype skype;
        public int sch = 0;
        public Form1()
        {
            InitializeComponent();
            skype = new Skype();
            skype.Attach(7, false);
            numericUpDown1.Maximum = skype.Chats.Count;
            //skype.MessageStatus+=new _ISkypeEvents_MessageStatusEventHandler(skype_MessageStatus);
        }


        void skype_MessageStatus(ChatMessage pMessage, TChatMessageStatus Status)
        {
            //if (pMessage.Body.IndexOf("!time") != -1)
            //skype.SendMessage(pMessage.Sender.Handle, (DateTime.Now.ToString()));
            sch++;
            if (sch >= 2)
            {
                if (pMessage.Chat.Members.Count <= 2)
                    skype.SendMessage(pMessage.Sender.Handle, "Привет :) \nЯ - бот, Антон ушел спать, он напишет тебе, как только проснется ;)");
                sch = 0;
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            richTextBox1.Text+="\n\n";
            SortedList<string, int> user_msgcount = new SortedList<string, int>();
            int edited = 0;
            int settopic = 0;
            int setpicture = 0;
            Int64 symbols = 0;
            //var chat = skype.RecentChats[(int)numericUpDown1.Value];
            //var chat = skype.get_Chat(textBox1.Text);
            //textBox1.Text = chat.Name;
            var messages = skype.get_Messages(textBox1.Text);
            for (int i = 1; i <= messages.Count; i++)
            {
                var msg = messages[i];

                if (user_msgcount.ContainsKey(msg.FromHandle))
                    user_msgcount[msg.FromHandle]++;
                else
                    user_msgcount.Add(msg.FromHandle, 1);


                if (msg.EditedBy != "")
                    edited++;
                if (msg.Type == TChatMessageType.cmeSetTopic)
                    settopic++;
                if (msg.Type == TChatMessageType.cmeSetPicture)
                    setpicture++;

                symbols += msg.Body.Length;
                Console.WriteLine(i.ToString());

            }
            //int msg_count = skype.RecentChats[(int)numericUpDown1.Value].Messages.Count;
            int msg_count = skype.get_Messages(textBox1.Text).Count;
            DateTime date_firts_msg = messages[messages.Count].Timestamp;
            DateTime date_last_msg = messages[1].Timestamp;
            //var sorteddict = user_msgcount.OrderBy(x => x.Value);

            //richTextBox1.Text += "////// Ktulhu Skype Bot \\\\\\\n";

            richTextBox1.Text += "Всего сообщений получено от "+textBox1.Text +" (во всех чатах) :" + msg_count.ToString() + "\n";
            richTextBox1.Text += "Из них отредактировано " + edited.ToString() + " сообщений\n";
            richTextBox1.Text += "Тему чата меняли " + settopic.ToString() + " раз\n";
            richTextBox1.Text += "Картинку чата меняли " + setpicture.ToString() + "раз\n";
            richTextBox1.Text += "Всего символов - " + symbols.ToString() + "\n";
            richTextBox1.Text += "Первое сообщ отправлено" + date_firts_msg.ToString() + "\n";
            richTextBox1.Text += "Последнее сообщ отправлено" + date_last_msg.ToString() + "\n";
            var sorted = user_msgcount.OrderBy(x => x.Value);

            foreach (var x in sorted)
                richTextBox1.Text += x.Key + " написал(а) " + x.Value + " cообщений\n";

            //richTextBox1.Text += "^^^^^^^^^ FHTAGH ^^^^^^^^";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //SortedList<string, int> my_chats = new SortedList<string, int>();
            //SortedList<string, DateTime> my_chats = new SortedList<string, DateTime>();


            int i = 0;
            new_data();
            foreach (Chat chat in skype.Chats)
            {


                if (chat.Members.Count == 2 && (chat.Members[1].Handle == textBox1.Text || chat.Members[2].Handle == textBox1.Text) && chat.Messages.Count != 0)
                {
                    if (i == 0)
                    {
                        Data.first_msg = chat.Messages[1].Timestamp;
                        Data.last_msg = chat.Messages[1].Timestamp;
                    }
                    //my_chats.Add(chat.Name, chat.Messages.Count);
                    //my_chats.Add(chat.Messages[1].Timestamp.ToString() +"      "+chat.Messages.Count+" "+ chat.Name, chat.Timestamp);

                    statistics(chat.Name);
                    i++;
                    //Console.WriteLine(i.ToString());
                }
            }
            Data.comp_name = textBox1.Text;
            //var sorted = my_chats.OrderBy(x => x.Value);

            //foreach (var x in sorted)
            //richTextBox1.Text += x.Value + "    " + x.Key + "\n";
            richTextBox1.Text = "";
            richTextBox1.Text += "<font size=\"15\">---------Ktulhu_Skype_Bot---------</font>" + "\n";
            richTextBox1.Text += "*********Статистика чата**********" + "\n";
            richTextBox1.Text += "*Всего сообщений отправлено - " + "<font color=\"#0000ff\"><u>" + (Data.comp_messages_count + Data.my_messages_count).ToString() + "</u></font>" + "\n";
            richTextBox1.Text += "**Из них " + "<font color=\"#00ff00\"><u>" + skype.CurrentUser.Handle + "</u></font>" + " отправил " + "<font color=\"#0000ff\"><u>" + Data.my_messages_count.ToString() + "</u></font>" + " сообщений" + "\n";
            richTextBox1.Text += "****Из них " + "<font color=\"#0000ff\"><u>" + Data.my_edited_count + "</u></font>" + " сообщений отредактировано" + "\n";
            richTextBox1.Text += "**Из них " + "<font color=\"#ff0000\"><u>" + Data.comp_name + "</u></font>" + " отправил(а) " + "<font color=\"#0000ff\"><u>" + Data.comp_messages_count.ToString() + "</u></font>" + " сообщений" + "\n";
            richTextBox1.Text += "****Из них " + "<font color=\"#0000ff\"><u>" + Data.comp_edited_count + "</u></font>" + " сообщений отредактировано" + "\n";
            richTextBox1.Text += "*" + "<font color=\"#00ff00\"><u>" + skype.CurrentUser.Handle + "</u></font>" + " отправил " + "<font color=\"#0000ff\"><u>" + Data.my_filetransfer_count.ToString() + "</u></font>" + " файлов" + "\n";
            richTextBox1.Text += "*" + "<font color=\"#ff0000\"><u>" + Data.comp_name + "</u></font>" + " отправил " + "<font color=\"#0000ff\"><u>" + Data.comp_filetransfer_count.ToString() + "</u></font>" + " файлов" + "\n";
            richTextBox1.Text += "*Всего символов отправлено " + "<font color=\"#0000ff\"><u>" + (Data.comp_symbols_count + Data.my_symbols_count).ToString() + "</u></font>" + " отправлено" + "\n";
            richTextBox1.Text += "** " + "<font color=\"#00ff00\"><u>" + skype.CurrentUser.Handle + "</u></font>" + " отправил " + "<font color=\"#0000ff\"><u>" + Data.my_symbols_count.ToString() + "</u></font>" + " символов" + "\n";
            richTextBox1.Text += "** " + "<font color=\"#ff0000\"><u>" + Data.comp_name + "</u></font>" + " отправил(a) " + "<font color=\"#0000ff\"><u>" + Data.comp_symbols_count.ToString() + "</u></font>" + " символов" + "\n";
            richTextBox1.Text += "Последнее сообщение отправлено " + "<font color=\"#0000ff\"><i>" + Data.first_msg.ToString() + "</i></font>" + "\n";
            richTextBox1.Text += "Первое сообщение отправлено " + "<font color=\"#0000ff\"><i>" + Data.last_msg.ToString() + "</i></font>" + "\n";
        }
        private void statistics(string chat_name)
        {
            var chat = skype.get_Chat(chat_name);
            var messages = chat.Messages;
            for (int i = 0; i < messages.Count; i++)
            {

                var msg = messages[i + 1];
                if (msg.FromHandle == skype.CurrentUser.Handle)
                {
                    Data.my_messages_count++;
                    Data.my_symbols_count += msg.Body.Length;
                    if (msg.EditedBy.Length != 0)
                        Data.my_edited_count++;
                    if (msg.Type == TChatMessageType.cmeEmoted)
                        Data.my_filetransfer_count++;

                }
                else
                {
                    Data.comp_messages_count++;
                    Data.comp_symbols_count += msg.Body.Length;
                    
                    if (msg.EditedBy.Length != 0)
                        Data.comp_edited_count++;
                    if (msg.Type == TChatMessageType.cmeEmoted)
                        Data.comp_filetransfer_count++;
                }
                Data.global_I++;
                Console.WriteLine(Data.global_I.ToString());
            }
            if (Data.last_msg > skype.get_Chat(chat_name).Messages[1].Timestamp)
                Data.last_msg = skype.get_Chat(chat_name).Messages[1].Timestamp;
            if (Data.first_msg < skype.get_Chat(chat_name).Messages[skype.get_Chat(chat_name).Messages.Count].Timestamp)
                Data.first_msg = skype.get_Chat(chat_name).Messages[skype.get_Chat(chat_name).Messages.Count].Timestamp;
        }

        private void new_data()
        {
            Data.comp_edited_count = 0;
            Data.comp_filetransfer_count = 0;
            Data.comp_messages_count = 0;
            Data.comp_name = "";
            Data.comp_symbols_count = 0;

            Data.my_edited_count = 0;
            Data.my_filetransfer_count = 0;
            Data.my_messages_count = 0;
            Data.my_symbols_count = 0;

            Data.first_msg = DateTime.Now;
            Data.last_msg = DateTime.Now;
            Data.global_I = 0;
        }



        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            var chat = skype.Chats[(int)numericUpDown1.Value];
            textBox1.Text = "";
            textBox1.Text += "сообщений - " + chat.Messages.Count + "/ учасников " + chat.Members.Count + "//" + chat.Name;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //foreach(Chat chat in skype.Chats)
            //{
            //    if (chat.Type == TChatType.chatTypeDialog || chat.Type == TChatType.chatTypeLegacyDialog)
            //        ;//if()
            //}
            SortedList<string, DateTime> my_chats = new SortedList<string, DateTime>();
            int i = 0;
            foreach (Chat chat in skype.Chats)
            {
                new_data();
                if (chat.Type == TChatType.chatTypeDialog || chat.Type == TChatType.chatTypeLegacyDialog)
                    //if (chat.Members.Count == 2 && (chat.Members[1].Handle == textBox1.Text || chat.Members[2].Handle == textBox1.Text) && chat.Messages.Count != 0)
                    if ((chat.Type == TChatType.chatTypeDialog) && chat.Members.Count != 1 &&(chat.Members[1].Handle == textBox1.Text || chat.Members[2].Handle == textBox1.Text) && chat.Messages.Count != 0)    
                {
                        my_chats.Add(chat.Name, chat.Messages[chat.Messages.Count].Timestamp);
                        //my_chats.Add(chat.Messages[1].Timestamp.ToString() +"      "+chat.Messages.Count+" "+ chat.Name, chat.Timestamp);

                        //statistics(chat.Name);
                        i++;
                        Console.WriteLine(i.ToString());
                    }
            }
            var sorted = my_chats.OrderBy(x => x.Value);
            foreach (var x in sorted)
                richTextBox1.Text += x.Value + "    " + x.Key + "\n";

        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            
            richTextBox1.Text = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var all_msg = "";
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Yellow;
            SortedList<string,USER_STAT> STAT = new SortedList<string,USER_STAT>();
            var messgages = skype.get_Chat("#daniil.koshelyuk/$3327b2e41ea6373f").Messages;
            int sch = 0;
            var msg_count = messgages.Count;
            
            foreach (ChatMessage msg in messgages)
            {
                if (!STAT.ContainsKey(msg.FromHandle))
                    STAT.Add(msg.FromHandle, new USER_STAT());


                var msg_body = msg.Body;
                all_msg += " " + msg_body;
                STAT[msg.FromHandle]._msg_count++;
                STAT[msg.FromHandle]._symb_count += msg.Body.Length;

                if (msg.EditedTimestamp.ToString() != "01.01.1970 4:00:00")
                    STAT[msg.FromHandle]._msg_edited_count++;

                if (msg.Type == TChatMessageType.cmeSetTopic)
                    STAT[msg.FromHandle]._topic_change_count++;

                if (msg.Type == TChatMessageType.cmeEmoted)
                    STAT[msg.FromHandle]._filetransg_count++;

                if (msg.Type == TChatMessageType.cmeSetPicture)
                    STAT[msg.FromHandle]._chatpict_change_count++;


                if (msg_body == "")
                    STAT[msg.FromHandle]._msg_delited_count++;

                while (msg_body.IndexOf("http://") != -1)
                {
                    STAT[msg.FromHandle]._http_count++;
                    msg_body = msg_body.Remove(msg_body.IndexOf("http://"), "http://".Length);

                }
                while (msg_body.IndexOf("(devil)") != -1)
                {
                    STAT[msg.FromHandle]._devilsmile++;
                    msg_body = msg_body.Remove(msg_body.IndexOf("(devil)"), "(devil)".Length);

                }
                int good = 0;
                int bad = 0;
               
                while (msg_body.IndexOf(")") != -1)
                {
                    //STAT[msg.FromHandle]._http_count++;
                    good++;
                    msg_body = msg_body.Remove(msg_body.IndexOf(")"), ")".Length);

                }
                while (msg_body.IndexOf("(") != -1)
                {
                    bad++;
                    msg_body = msg_body.Remove(msg_body.IndexOf("("), "(".Length);

                }
                

                if (good - bad > 0)
                    STAT[msg.FromHandle]._good_smile += (good - bad);
                else
                    STAT[msg.FromHandle]._bad_smile += (bad - good);

               
                //richTextBox1.Text += sch.ToString() + "\n";
                //richTextBox1.SelectionStart = richTextBox1.Text.Length - 1;
                //richTextBox1.Refresh();
                Console.WriteLine(sch.ToString());
                sch++;
                //panel1.BringToFront();
                //progressBar1.BringToFront();
                progressBar1.Value = (sch)*100/msg_count;
                progressBar1.Refresh();
                
                
            }
            foreach (var item in STAT)
            {
                listView1.Items.Add(new ListViewItem(new string[] { item.Key, item.Value._msg_count.ToString(), item.Value._msg_edited_count.ToString(), item.Value._msg_delited_count.ToString(), item.Value._symb_count.ToString(), item.Value._filetransg_count.ToString(), item.Value._topic_change_count.ToString(), item.Value._chatpict_change_count.ToString(), item.Value._http_count.ToString(), item.Value._devilsmile.ToString(), item.Value._good_smile.ToString(), item.Value._bad_smile.ToString(), item.Value._xerr.ToString() }));
            }
          
            words(all_msg);

        }
        private void words(string str)
        {
            //var messgages = skype.get_Chat("#daniil.koshelyuk/$3327b2e41ea6373f").Messages;
            SortedList<string, int> words = new SortedList<string, int>();

            str = str.ToLower().Replace("\n", " ").Replace("\r", null);
            while (str.Contains(@"  "))
                str = str.Replace(@"  ", @" ");

            var count = 0;
            while(count!=str.Length)
            {
                if(str[count]==' ')
                    str = str.Remove(count, 1+count);

                var nomer = str.IndexOf(@" ",count);
                if (nomer==-1)
                    break;
                    
                if (words.ContainsKey(str.Substring(count, nomer-count)))
                    words[str.Substring(count, nomer-count)]++;
                else
                    words.Add(str.Substring(count, nomer-count),1);
                count = nomer+1;
                //Console.WriteLine(count.ToString());
            }
            foreach(var item in words)
            {
                listView2.Items.Add(new ListViewItem(new string[] {item.Key, item.Value.ToString()}));
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ListView lv = (ListView)sender;
            lv.BeginUpdate();
            lv.ListViewItemSorter = new ListViewItemComparer(e.Column, 1);
            lv.EndUpdate();
        }

        private void listView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ListView lv = (ListView)sender;
            lv.BeginUpdate();
            lv.ListViewItemSorter = new ListViewItemComparer(e.Column, 1);
            lv.EndUpdate();
        }
    }
}
