using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using AForge;
using AForge.Controls;
using AForge.Imaging;
using AForge.Math;
using AForge.Video;
using AForge.Video.DirectShow;
using Microsoft;
using Microsoft.ProjectOxford;
using Microsoft.ProjectOxford.Common;
using Microsoft.ProjectOxford.Emotion;
using Microsoft.ProjectOxford.Emotion.Contract;
using System.Drawing.Imaging;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Media;
using System.Diagnostics;

namespace Hachathon_Project2
{
    public partial class Form1 : Form
    {
        int flag = 0;
        int pretm;
        int prefeal=-1;
        public Form1()
        {
            InitializeComponent();
            flag = 0;
            pretm = 0;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            if (flag == 1) return;
            flag = 1;
            while(true)
            {
                    await RUN();
                for(int i=0;i<20;i++)
                    System.Threading.Thread.Sleep(100);
            }
            
        }
        public async Task<bool> RUN()
        {
            string imagename = @"MyEmotion.jpg";
            string subscriptionKey = "8d3ba73818224d5b8ead829fa45279be";
            CaptureImage(imagename);
           // MessageBox.Show("fff");
            FileStream file = new FileStream(imagename, FileMode.OpenOrCreate);
            EmotionServiceClient client = new EmotionServiceClient(subscriptionKey);
            Microsoft.ProjectOxford.Emotion.Contract.Emotion[] emotionress;
            //MessageBox.Show("ffffff");
            try
            {
                labelOutput.Text = "Connecting! Wait for response...";
                emotionress = await client.RecognizeAsync(file);
                //MessageBox.Show("ffffff");
                labelOutput.Text = "Connected!";
                file.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                file.Close();
                return false;
            }
            //MessageBox.Show("nihao");
            int feal = Getfeal(emotionress);
            //MessageBox.Show("Emotion: "+feal.ToString());
            if (feal != prefeal)
            {
                ChangeImage(feal);
                ChangeMusic(feal);
                ChangeExe(feal);
                prefeal = feal;
            }

            file.Close();
            labelOutput.Text = "Done!Emotion:"+feal.ToString();
            flag = 0;
            return true;
        }
        private FilterInfoCollection videoDevices;
        private VideoCaptureDevice videoSource;
        public int selectedDeviceIndex = 0;
        public void CaptureImage(string imagename)//抓取表情
        {
           // MessageBox.Show("Here!");
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            selectedDeviceIndex = 0;
            videoSource = new VideoCaptureDevice(videoDevices[selectedDeviceIndex].MonikerString);//连接摄像头。
            videoSource.VideoResolution = videoSource.VideoCapabilities[selectedDeviceIndex];
            videoSourcePlayer1.VideoSource = videoSource;
            // set NewFrame event handler
            videoSourcePlayer1.Start();
            if (videoSource == null)
                return;
            // videoSourcePlayer1.
            //Console.WriteLine();
            //MessageBox.Show("Here!");
            System.Threading.Thread.Sleep(2000);
            Bitmap bitmap = null;
            bitmap = videoSourcePlayer1.GetCurrentVideoFrame();
            //videoSourcePlayer1.Show();
            if (bitmap == null)
                return;
           // MessageBox.Show("Here!");
            string fileName = imagename;
            //MessageBox.Show(fileName);
            bitmap.Save(fileName, ImageFormat.Jpeg);
            bitmap.Dispose();
            bitmap = null;
           // MessageBox.Show("here");
            videoSourcePlayer1.SignalToStop();
            videoSourcePlayer1.WaitForStop();
        }
        public class WinAPI
        {
            [DllImport("user32.dll", EntryPoint = "SystemParametersInfo")]
            public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }
        public void ChangeImage(int feal)//切换壁纸
        {
            
            if (feal < 0) return;
            string filename;
            filename = @"C:\Users\reminiscence\Pictures\壁纸\壁纸\" + feal.ToString() + "\\0.jpg";
            //MessageBox.Show(filename);
            string bmpPath = filename;//新图片要存储的位置
            string ImagePath = filename;//图片路径
            string ImageFormt = ImagePath.Substring(ImagePath.LastIndexOf('.'));//获取图片格式 
            //MessageBox.Show(File.Exists(bmpPath).ToString());
            int nResult;
            if (File.Exists(bmpPath))
            {

                nResult = WinAPI.SystemParametersInfo(20, 1, bmpPath, 0x1 | 0x2); //更换壁纸
                if (nResult == 0)
                {
                    MessageBox.Show("没有更新成功!");
                }
                else
                {
                    RegistryKey hk = Registry.CurrentUser;
                    RegistryKey run = hk.CreateSubKey(@"Control Panel\Desktop\");
                    run.SetValue("Wallpaper", bmpPath); //将新图片路径写入注册表
                                                        //MessageBox.Show("更新成功");
                }
            }
            else
            {
                MessageBox.Show(bmpPath + "文件不存在!");
            }
        }
        public void ChangeMusic(int feal)//切换音乐
        {
            if (feal < 0) return;
            System.Media.SoundPlayer sp = new SoundPlayer();
            sp.SoundLocation = @"C:\Users\reminiscence\Pictures\壁纸\壁纸\" + feal.ToString()+@"\0.wav";
            sp.PlayLooping();
        }
        public void ChangeExe(int feal)
        {
            if (feal == 4)
            {
                //Get path of the system folder.
                string sysFolder =
                Environment.GetFolderPath(Environment.SpecialFolder.System);
                //Create a new ProcessStartInfo structure.
                ProcessStartInfo pInfo = new ProcessStartInfo();
                //Set the file name member. 
                // pInfo.FileName = sysFolder + @"\123.txt";
                pInfo.FileName = @"D:\CodeBlocks\codeblocks.exe";
                //UseShellExecute is true by default. It is set here for illustration.
                pInfo.UseShellExecute = true;
                Process p = Process.Start(pInfo);
            }
            else if(feal==5)
            {
                //Get path of the system folder.
                string sysFolder =
                Environment.GetFolderPath(Environment.SpecialFolder.System);
                //Create a new ProcessStartInfo structure.
                ProcessStartInfo pInfo = new ProcessStartInfo();
                //Set the file name member. 
                // pInfo.FileName = sysFolder + @"\123.txt";
                pInfo.FileName = @"D:\Program Files (x86)\JetBrains\WebStorm 2016.2\bin\WebStorm.exe";
                //UseShellExecute is true by default. It is set here for illustration.
                pInfo.UseShellExecute = true;
                Process p = Process.Start(pInfo);
            }
        }
        public int Getfeal(Emotion[] emotionres)
        {
            
            if (emotionres.Length == 0)
            {
            //    MessageBox.Show("hh");
                return -1;
            }
            float[] emotionvalue = new float[10];
            int i = 0;
            for (i = 0; i < 10; i++)
            {
                emotionvalue[i] = 0.0F;
            }
            for (i = 0; i < emotionres.Length; i++)
            {
                emotionvalue[0] += emotionres[i].Scores.Anger;
                emotionvalue[1] += emotionres[i].Scores.Contempt;
                emotionvalue[2] += emotionres[i].Scores.Disgust;
                emotionvalue[3] += emotionres[i].Scores.Fear;
                emotionvalue[4] += emotionres[i].Scores.Happiness;
                emotionvalue[5] += emotionres[i].Scores.Neutral;
                emotionvalue[6] += emotionres[i].Scores.Sadness;
                emotionvalue[7] += emotionres[i].Scores.Surprise;
            }
            for (i = 0; i < 10; i++)
            {
                emotionvalue[i] /= emotionres.Length;
            }
            i = 0;
            int tmpres = 0;
            for (i = 1; i < 10; i++)
            {
                if (emotionvalue[i] > emotionvalue[tmpres])
                {
                    tmpres = i;
                }
            }
            //MessageBox.Show(tmpres.ToString()+":"+emotionvalue[tmpres].ToString());
            return tmpres;
        }

        private void labelOutput_Click(object sender, EventArgs e)
        {

        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            Button button1 = sender as Button;
            button1.FlatAppearance.BorderSize = 1;
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            Button button1 = sender as Button;
            button1.FlatAppearance.BorderSize=0;
        }

        private async void button1_Click_1(object sender, EventArgs e)
        {
            if (flag == 1) return;
            flag = 1;
            while (true)
            {
                await RUN();
                System.Threading.Thread.Sleep(7000);
            }
        }
    }
}
