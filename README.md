# MyArp (version 1.2, released December 5, 2021)

MyArp is a free open source alternative command line program to Microsoft's Arp program. 

MyArp is licensed under the [MIT license](https://rlatour.com/myarp/license.txt).

**Some background:**

ARP stands for Address Resolution Protocol, to learn more you can read all about it [here](https://en.wikipedia.org/wiki/Address_Resolution_Protocol)

Microsoft's Arp program is documented [here](https://technet.microsoft.com/en-us/library/cc754761(v=ws.11).aspx)

**What MyArp does:**

MyArp, like ARP, is a command line program.  It does not do all the things APR does, but for what it does - it does more.

MyArp is focused on reporting devices that are either connected or have been connected to your network.  

Unlike Microsoft's ARP program, it is does not add to or delete from the ARP cache.

Below are two screenshots, the one on the taken from Microsoft's ARP, and the other from MyARP:

**Microsoft's ARP**

![Microsoft's Arp](/images/arp.jpg)

**MyArp**

![Microsoft's Arp]/images/myarp.jpg)

**MyArp provides the following features over ARP:**

\- reports device names  
\- allows you to add and edit user friendly device descriptions that are also reported  
\- keeps a history of devices that have been seen in the past, and if they are not actively connected to your network reports when they were last seen  
\- provides more accurate reporting by pinging devices on your network before reporting them (as this can be time consuming there is also an option to not ping)

**Here is how you can use MyArp from the command line**:
  
MyArp /? /ADD \[Physical Address\] (Description) /DEL \[Physical Address\] /C /DBB /DBD /DBE /DBR /NP /NRA /NRI /NRD /NRS /P /Q /R

/? = show (this) help and exit  
  
/ADD \[Physical address\] (Description) = Add or update a database entry and description  
   only one /ADD statement is allowed at a time  
   an /ADD statement must be the only statement on a line  
   example /ADD statements look like this:  
      /ADD 7C:DD:90:00:00:01  
      /ADD 7C:DD:90:00:00:02 Raspberry Pi Wireless  
  
/DEL \[Physical address\] = Delete a database entry  
   only one /DEL statement is allowed at a time  
   a /DEL statement must be the only statement on a line  
   an example /DEL statement looks like this:  
      /DEL 7C:DD:90:00:00:01  
  
/C = clear console window before writing report  
  
/DBB = database backup  
/DBD = database delete  
/DBE = database edit  
/DBR = database restore from backup  
  
/HYPERV = consolidate Hyper-V related entries as 234.-.-.-  
  
/NP = do not ping (saves time but results may be less accurate)  
  
/NRA = do not report active devices  
/NRI = do not report inactive devices  
/NRD = do not report dynamic IP addresses  
/NRS = do not report static IP addresses  
  
/P = pause prompt before exit  
/Q = no prompts or reports (just update the database)  
/R = refresh device names (may take a long time)

**Here is how you can use MyArp from a desktop icon**

While MyArp is a command line program, it can still be run by via a desktop icon.  Here is how you can set the program up as a desktop icon (please note the arrow points to where the command line parameters are added):

![MyArp desktop icon](/images/desktopicon.jpg)

# Download

You are welcome to download and use MyArp for free on as many computers as you would like.

A downloadable signed executable version of the program is available from [here](https://github.com/roblatour/myarp/releases/download/v1.2.0.0/myarpsetup.exe).

* * *
 ## Supporting MyArp

To help support MyArp, or to just say thanks, you're welcome to 'buy me a coffee'<br><br>
[<img alt="buy me  a coffee" width="200px" src="https://cdn.buymeacoffee.com/buttons/v2/default-blue.png" />](https://www.buymeacoffee.com/roblatour)
* * *
Copyright © 2021 Rob Latour
* * *   
