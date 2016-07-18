# What is BlueSol_NET?
A set of wrapper modules and an example program for use with the Blue Soleil SDK, to demonstrate how to interface with each Bluetooth profile.  Most of the profiles are already supported. Only tested under x86 compiling due to heavy use of function pointers and data-type conversions.  All marshalling is done by hand, in code, building/parsing structures out of byte arrays.  Almost every callback is implemented as a .Net event.  Almost every function is wrapped in a .Net friendly manner.  

# What is Blue Soleil?
Blue Soleil is a third-party bluetooth library and bluetooth stack.  It comes with an SDK and example code written in C, and the progamming interface allows audio connections (which no other free/inexpensive bluetooth stack provides).  Their support is limited, and their code/examples are very difficult to use.  An incorrect declaration or improper use of a function can lead to memory corruption, things not working, and ultimately crashing.  And none of it is thread-safe.  So most developers spend hours (weeks, months) just trying to get the examples translated.  BlueSol_NET is my attempt at translating and wrapping their SDK into a reusable library.


# Bluetooth Profiles Supported:

<b>-PBAP - PhoneBook Access Profile - 100% complete.  </b>
Download VCards of contacts, with pictures.  VCard (.vcf) parser included to read the info and image of each contact.

<b>-MAP - Message Access Profile - 95% complete.</b>
Send messages.  Retrieve messages.  
The only thing that doesn't work is the "new message" event.  So you have to poll the device for new messages.
(note: iPhone's do not allow sending of text messages via bluetooth MAP.)

<b>-OBEX - Object Exchange Profile - 100% complete. </b>
This is used to transfer add a contact to the phone.  Use the included VCard module to compose a VCard file (specifying contact info, picture, phone numbers, addresses) and push it to the phone.

<b>-PAN - Personal Area Network - 100% complete.</b>
This is used to tether your PC or tablet to your phone and share your phone's internet connection.  Must be enabled on the phone.

<b>-HFP - Hands Free Profile - 99% complete.</b>
Make and receive phone calls.  Includes cellular network information, Caller-ID, transferring audio to/from phone, sending DTMF tones, etc.

<b>-AVRCP - Audio/Video Remote-Control Profile - 90% complete.</b>
Play music from your phone on your PC.  Just like a bluetooth head unit in a car.
The only thing that doesn't work is media browsing, which seems to be a new/advanced feature not supported by all phones.

<b>-A2DP - Advanced Audio Distribution Profile - 95% complete.</b>
This is the underlying audio connection used for both phone calls (HFP) and media playback (AVRCP).  

<b>-FTP - File Transfer Profile - 98% complete. </b>
Browse the folders and files on the phone.  Transfer files.  Delete files.

<b>-SPP - Serial Port Profile - 70% complete.</b>
Enable and connect to a serial port on a remote device, such as a phone or an OBD reader, or any BT device that exposes a serial port.
Once the port is connected, standard Windows serial communication can be used.

<b>-HID - Human Input Device - 10% complete.</b>
I have no sample code from the SDK, and no HID devices to test with.  So only some of the declarations are complete.

# Screenshots
![BlueSolNet ScreenShot](http://www.compulsivecode.com/images/bluesoltest_ss.png "BlueSolNet ScreenShot")



