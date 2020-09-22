<div align="center">

## Msn 6\.2 API Help


</div>

### Description

This article will explain a much more better example of Msn 6 API. From what I've seen, too many people submit windows messenger code, not msn 6.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chaositic Serge](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chaositic-serge.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chaositic-serge-msn-6-2-api-help__1-55522/archive/master.zip)





### Source Code

```
This is my first tutorial so bare with me. In an article that explains the basics of Msn 6 API, the author only gave you the references and declarations and what to put in form_load(). This article will expand on that. Basically, when you have referenced
Messenger API Type Library, have put:
Public WithEvents Msn as Messenger in Declarations, and have put:
Private Sub Form_Load()
Set Msn = New Messenger
End Sub
you would tend to wonder how to do all the commands properly.
In that one article that gave the basics, the author asked for a comment of how to change your nickname. But before I give the code, you have to put the following into a module:
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Now, here is the discovered nickname changing code:
Private Sub Command1_Click()
msn.OptionsPages 0, MOPT_GENERAL_PAGE
SendKeys Text1.Text
sleep (1000)
SendKeys "{ENTER}"
End Sub
To make proper use of the code of'course, you have to make the command button and the textbox control. The rest is quite simple.
Now to explain this code a bit.
msn.OptionsPages 0, MOPT_GENERAL_PAGE
This code bit means it will open up the general page located in your personal settings where you change your name. In windows messenger, you were able to change your nickname using one line of code. But Msn 6.2 is much different so you have to bare with the changes.
SendKeys Text1.Text
This bit of the code uses Sendkeys which allows you to force certain keys pressed. Anything in text1.text while pressing the command button will be sent.
sleep (1000)
This code tidbit performs the sleep command which you earlier allowed in the module. The 1000 is the amount of milliseconds. 1000 = 1 second and this is required for the program to change the name properly. You are free to edit this number in additions of 1000 if 1000 isn't enough.
SendKeys "{ENTER}"
The last part of the code. This just forces the pressing of Enter key. When the personal settings opens and goes to general page, when you press enter, it automatically clicks ok. Thats all this code does, press enter which clicks the ok button.
Now run it and test it. If it works, great. You're doing well. If not, you didn't do something correctly.
Now, moving on. You've mastered an a new way of changing your nickname but what's next? Well, one thing to discover is changing your status. It's quite simple. If you create another button and put its caption as online and in the code put:
Private Sub Command2_Click()
msn.mystatus = mistatus_online
End Sub
It would change your status to Online if you ran the program and clicked the button. Now that that's done, you can easily figure out the rest of the statuses. Observe.
msn.mystatus = mistatus_offline = signed out
msn.MyStatus = MISTATUS_BUSY = busy status
msn.MyStatus = MISTATUS_Be_Right_Back = brb
msn.MyStatus = mistatus_away = away
msn.MyStatus = MISTATUS_INVISIBLE = appear offline
You get the point.
msn.MyStatus = MISTATUS_ON_THE_PHONE
msn.MyStatus = MISTATUS_OUT_TO_LUNCH
Now, you know how to change your status. It was pretty easy right?
Ok, now we try some more code and then I'm out for now.
This is a trick I found out not too long ago.
Msn.autosignin
Notice when you type it that nothing appears beside it for adding in variables etc.
For this to work, you would create a timer control and it would look like this:
Private Sub Timer1_Timer()
if msn.mystatus = mistatus_offline = true then
msn.autosignin
end if
End Sub
Set the timers interval to 10. Basically, this trick checks to see if you have signed out. Like let's say you've been disconnected. With this program running, it would auto sign you back in immediately. Even though I think that when you get disconnected from the net, it auto signs you back in for you but, it can also be useful to auto sign you in when someone signs you out.
Anyways, I hope you found this tutorial helpful. One last note before I go: This only works on XP because I program on XP and thats all i've tested it on. So far, anyone who has windows messenger can program msn 6.2 properly. I know there are ways to program in windows 98 but I don't know how so sorry. I hope the tutorial gave you an idea of API more for msn. Please give feedback and any questions, i'll gladly answer!
```

