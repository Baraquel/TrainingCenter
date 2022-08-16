# TrainingCenter  

## Microsoft courses deployement script  
    Author: Mosselmans Benjamin  
    Linkedin link: https://www.linkedin.com/in/mosselmansben/  
    Date: Summer 2022  
    Function : Seek and deploy  
    Description: This script receive as parameters the id (int) of the Microsoft course and if a boolean of the computer state (Microsoft trainer computer). With these       two parameters, it'll seek on a share  
    on the network (Here: \\tke-veeam\E\MOC) for the folder containing such id, get all txt extension file names in the folder containing the Virtual Machines to then       find on the other directory containing all  
    base drives the corresponding base drives necessary to said course. It'll then execute the unzipping executable linked to the bases. When done, it executes the           scripts given by Microsoft to deploy and snapshots the created Virtual Machines.  
    If the computer is tagged as teacher, it'll also deploy the powerpoint and onenote of the course.  
    The second part of the script go inside each Virtual Machines, and if given the right circunstences (Remote allowed), it'll delete in the registry keyboard registry     keys to only let English and French(Belgium).  
    The IME aren't deleted.  
## Parameters:  
- [String] Course's id
- [Boolean] Computer state (Microsoft Trainer computer)  
## Possible upgrades:  
- Stock all Microsoft administrator credentials in a text file to be feed to the code with iteration until it finds the right credential to use.
- Change the SendKeys function to one allowed by MDT directly injected in the task sequence (System.Windows.Forms.SendKeys function not working                             through MDT)
- Delete the additionnal IME
- Silent rearm (slmgr -rearm)
