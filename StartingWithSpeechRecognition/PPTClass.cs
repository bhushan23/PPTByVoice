using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ppt = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
namespace StartingWithSpeechRecognition
{
    class PPTClass
    {
        public int total_slides =0,currentslide=0;
        ppt._Application pApp;
        ppt.Presentation pPre;
        public PPTClass()
        { 
        
        }
        public PPTClass(string pptfile)
        {
            pApp = new ppt.Application();
            pPre = pApp.Presentations.Open(pptfile, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue);
            pPre.SlideShowSettings.Run();
         
            total_slides = pPre.Slides.Count;
            currentslide = 1;
        }
        public void next_slide()
        {
            try
            {
                pPre.SlideShowWindow.View.Next();
            }
            catch (Exception ex)
            {
               
                pPre.SlideShowSettings.Run();
               
               pPre.SlideShowWindow.View.GotoSlide(currentslide);
                //pPre.SlideShowWindow.View.Next();

            }
            currentslide++;
   
        }
        public void prev_slide()
        {
            try{
            pPre.SlideShowWindow.View.Previous();
            }     catch (Exception ex1)
            {
               
                pPre.SlideShowSettings.Run();
                pPre.SlideShowWindow.View.GotoSlide(currentslide);
           

            }
            currentslide--;
        }
        public int can_recognizer_continue()
        {
          
            if (currentslide <= total_slides)
                return 1;
            else
                return 0;
        }
        public void close_ppt()
        {
            pPre.Close();

        }
    }
}
