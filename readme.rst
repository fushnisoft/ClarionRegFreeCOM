=============================
Clarion Registration-Free COM
=============================

This project is a demonstration of my attempts to use Registration-Free Activation of COM Components with Clarion.

- `Registration-Free Activation of .NET-Based Components: A Walkthrough <http://msdn.microsoft.com/en-us/library/ms973915.aspx>`_.
- `Getting COM registration data with the activation context API <http://www.voyce.com/index.php/2007/08/15/getting-com-registration-data-with-the-activation-context-api/>`_.
- `RegFreeCom tag on StackOverflow <http://stackoverflow.com/questions/tagged/regfreecom>`_.

Installation
============

You can download the files as a zip file from the link on this page or better yet, go get the awesome `GitHub windows client <http://windows.github.com/>`_ !

I have included the binary files so you can just run the InteropTest.exe and see what happens. If you are interested in taking a look and compiling it for yourself I have included both the Visual Studio 2010 C# project and the Clarion8 project (source only, no app or templates required). Have a look in the \Src directory, it should be fairly self explanatory.

What is happening?
==================

Once you have the manifest files correct you need to use the cslid instead of the progid in myFeq{PROP:CREATE} and calling methods in the object seems to be working fine.

However, this only gives you an ole object (myFeq{PROP:OBJECT} returns a value like `1234) but myFeq{PROP:OLE} is FALSE which based on the docs indicates that there is no object in the container.

Is there anyway to take the object id I have and activate it (display it?) in the container?

Am I in fact barking up the wrong tree with this?! Not sure, I am learning :)

Here are some links to things I intend to try next:

- Hosting the CLR directly and loading the .Net control that way might be the way to go. This looks like a pretty complex project so for now it will have to wait but I think it will be the best solution in the long run if it works. `Starting an instance on NET from Clarion using COM <http://www.icetips.com/showarticle.php?articleid=304>`_ looks like a good start although it will need to be updated for .NET v4 framework,
- This looks like it is about creating a control without using the clarion OLE container directly. If it works, it may use the manifest correctly and would be a nice start: http://www.clarionopensource.com/ClarionCOM/DiYIEExample.htm
- There are a bunch of other COM articles over at IceTips that I had not seen until now that look very helpful: http://www.icetips.com/articleindex.php#category10

Let me know how you go!