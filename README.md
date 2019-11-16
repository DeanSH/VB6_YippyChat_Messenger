# VB6_YippyChat_Messenger
YippyChat Messenger was a Windows-based application made using VB6 that very much replicated Yahoo Messenger after they closed chat rooms in 2012, but had secure data protocol encryption and newer features than Yahoo such as Screen sharing and chat room Tattoo's. I owned and operated YippyChat successfully for 2 years before deciding to close it down, it was comprised of multiple server-side application to fully function & load balance. 

YippyChat supported such features as Registration, Login, preferences, ignores, private messaging, skins, profile pictures, friends list, custom status messages, chat rooms with voice channels, smileys, webcams, voice calling, file transfers and much more!

# YppyChat Messenger Client Side Application Notes:
"YippChatClient.zip" contains the actually client-side YippyChat Messengers VB6 source code. There are multiple sub folders with it some of which I had to compress the contents for into a yet another Zip, such as Flags, Skins and Controls! So be sure to find those folders and Extract the additional Zip folders inside them! There is also an important Avatars folder which held over a thousand avatar images, I replaced it with a Read Me basically informing you to find the Avatars.zip file seperately on this GitHub repo and to move it into that Avatar folder and extract it there!

Once you have done all these extractions it should now be ready to run the "Run-Time-Files.exe" which will register all the needed files for this project, followed by installing an important Voice-Fix and TSP Codec installation which ensures that voice communications will be fully working in the project and for users of the software if compiled and distributed, these Controls are vital. If you see any Read Me files, then its probably best to also read them! The Extensions folder is for supporting Cam and Screen sharing which are seperate extension exe files essentially that execute when required by the main YippyChat Messenger.

# "Reg Files YippyChat.zip" Notes:
"Reg Files YippyChat.zip" contains the all the needed files, and installers for YippyChat and Voice communications along with the actually Source Code for creating the "Run-Time-Files.exe" incase any modification to these might be required.

# "YippyVoice32 dll" Notes:
"YippyVoice32 dll" folder contains the Source Code associated with the actual YippyVoice32.dll that gets registered to enable recording microphone audio for streaming and handles the playback of recieved audio.

# "YippyChat Cam Or Screen Sharing" Notes:
"YippyChat Cam Or Screen Sharing" folder contains the source code examples for the Cam and Screen sharing extension applications!

# YippyChat Responsive Static Website Notes:
"YippyChat Responsive Static Website" folder contains all the Web Files that previously provided a advertising and download point for YippChat Messenger at the old domain yippychat.com, there is also VERY IMPORTANT folders and scripts in here that help YippyChat Messenger to actually work for certain functionality such as Profile Pictures uploading and display cabalities as they where stored in the Web hosting via the uploads folder, and the profilepic.php or getfile.php scripts!! The adverts.php script helps support display of Sponsors in YippyChat Messenger and update.php contains the current version details to inform users if there was an update available!

The getfile.php code follows best practice for allowing image uploads to the Web hosting by whitelisting 'jpg', 'png', 'gif', 'jpeg', 'bmp', 'JPG', 'JPEG' filetypes, blocking 'php', 'php3', 'php4', 'phtml','exe' file types, and restricting max file name lengths, file name characters, file sizes upto 5mb, blank file names, validates MIME image data type and run images through a get image file size function, double check the MIME data a second way, checks if the profile picture for that user already exists to over write it, removing it first, renames the image file before saving to prevent hackers find the file name in the web hosting and trying to trick the server into executing a malicous file somehow, verifies that the file was successfully uploaded with its secret file naming, then for futher security resizes the image to save space, further validating it actually an image and erasing any malicous code potentially attached, then lastly also forces it to become a JPG image file type regardless of the type uploaded.

There is 2 .htaccess files which are vital to support the Messenger too, they mask the file paths with redirects, one is in the root folder, and the other in the upload folder where display pictures go which limits file types that can execute in that folder!

# YippyChat Messenger Server-Side Application Notes:
These Server-Side Applications are Windows-based exes when compiled and designed to run on dedicated servers, running multipe instances of Windows VPS's which I personally used V-Sphere and ESXI 4.0 for originally for the virtualisation, each dedicated server required its own static IP address, however that could be mitigated using Domain DNS records. If setting up these Servers they must be configured correctly to point at the Main Server or Main Voice server, etc, and YippyChat Messengers connection for login and other features in the Client-Side application would all require updating to point to everything perfectly!

MAIN, REGISTER & LOGIN SERVERS!!

-> "Main Server and Reg Server" folder contains 2 project source codes, the Reg server handles New Account registration and needs to be located in the same folder as the Main Server on a single VPS because they share the same Activity Log folder which acts as the database for YippyChat Messenger, where a folder fo each new account is created and a few txt files are stored containing various informations such as account details, ignores, offline messages recieved, etc. The Main Server application is the most important for YippyChat Messenger to be able to work, it acts as the centralised hub which Pre-Login and Sub-Login servers connect with to relay data between multiple Sub server instances across multiple VPS's and/or dedicated servers, handling load balancing between them all too! Basically when someone logs in they connect to 1 of multiple Pre-Login server instances which will then communicate with the Main server and ask which Sub-Login server has the least connections, the Main server asks connected Sub-Login servers for connection counts, and determines which is best suited returning the appropriate Sub-Login server that is online to be used! The user is then directed to that Sub-Login server to complete login, which validates the password by requesting the Main Server to confirm it, and then Users remain Connect to this Sub-Login server which enables them to do activities like, add friends, private messaging and signals many other things. Note that all data between the Client YippyChat Messenger and Server-Side Apps gets Encrypted uniquely.

-> "Pre Login Server"	contains the Source Code for the Pre-Login server described above with the Main Server flow description!
-> "Sub Login Server"	contains the Source Code for the Sub-Login server described above with the Main Server flow description! The Sub-Login servers also send regular Pings to all their connected users in order to keep the socket connection alive and detect dead connections, setting the User to offline again, this could occur if they did not logout but suddenly DC from the internet!

RELAY SERVERS!!

-> "Call Server" folder contians the Source Code for the calling middle man server that relays data between peers in a call after singalling through the Sub-Login & Main Server is successful, as in a user accepts the incoming call from another user!
-> "Cam Server"	folder contians the Source Code for the Cam share middle man server that relays data between peers in a Cam session after singalling through the Sub-Login & Main Server is successful, as in a user accepts the incoming Cam invitation from another user!
-> "Desktop Server"	folder contians the Source Code for the Screen share middle man server that relays data between peers after singalling through the Sub-Login & Main Server is successful, as in a user accepts the incoming Screen share invitation from a user!
-> "File Server" folder contians the Source Code for the File transfer middle man server that relays file data between peers after singalling through the Sub-Login & Main Server is successful, as in a user accepts the incoming File Transfer from a user!

RELAY SERVERS WHERE CREATED OPPOSED TO P2P IN ORDER TO AVOID NAT TRAVERSAL ISSUES, OR USER IP SNIFFING!

CHAT SERVER!!

-> "Chat Server" folder contains the Source Code for the Chat rooms server which can run multiple instances at a time in a single VPS node, and across multiple VPS nodes and needs to be configured correctly, no 2 instance on a single VPS node should use the same chat ports, they must listen on different ports, this Chat server handles who is currently in which room, limits how many per room, a single instance manages multiple rooms at once that you load into it, I personally only ran 15 rooms per instance. There is also Server-Side Admin controls here allowing to block or kick users if required from a chat room.

VOICE SERVERS!!

-> "Main Voice Server" folder contains the Source Code for the Main Voice server, This acts as a centralised hub for the Sub Voice Servers and allows Chatters to connect and discover the appropriate Voice Channel IP adn Port configuration to join the right Sub Voice Server that has the channel relating to the room that user has joined!

-> "Sub Voice Server" folder contains the Source Code for the Sub Voice server which can run multiple instances at a time in a single VPS node, and across multiple VPS nodes and needs to be configured correctly to point at the Main Voice Server, and no 2 instances on a single VPS node should use the same voice ports, they must listen on different ports, this Chat server handles who is currently in which rooms voice channel, limits how many per voice channel, a single instance manages multiple rooms voice connectivity at once according to the room names that you load into it, I personally only ran 15 rooms per instance to match the Chat Server instances.

