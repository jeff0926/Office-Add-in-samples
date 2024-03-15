MS office Add-In:

<https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/word-tutorial>

prerequirements:

make sure you have yo generator

\`\`\`npm install -g yo generator-office

\`\`\`mkdir create folder \> udex-office-add-in

![](media/image1.png){width="6.5in" height="1.2965277777777777in"}

\`\`\`cd \\udex-office-add-in

\`\`\`yo office

![](media/image2.png){width="6.30129593175853in"
height="4.7702373140857395in"}

![](media/image3.png){width="6.5in" height="3.4402777777777778in"}

![](media/image4.png){width="6.5in" height="3.3743055555555554in"}

![](media/image5.png){width="6.5in" height="3.84375in"}

![](media/image6.png){width="6.259633639545057in"
height="4.822314085739283in"}

![](media/image7.png){width="6.5in" height="2.5305555555555554in"}

Navigate to folder :

Cd \\udex-office-add-in\\udex-author-assist\\src\\taskpane

![](media/image8.png){width="6.5in" height="2.1222222222222222in"}

\<Code \> \<button class=\"ms-Button\" id=\"insert-paragraph\"\>Insert
Paragraph\</button\>\<br/\>\<br/\>

![](media/image9.png){width="6.5in" height="1.5840277777777778in"}

Open: taskpane.js

\`\`\`

document.getElementById(\"insert-paragraph\").onclick = () =\>
tryCatch(insertParagraph);

![](media/image10.png){width="6.5in" height="2.0993055555555555in"}

![](media/image11.png){width="6.5in" height="0.8736111111111111in"}

**test your add-in in Word,** run the following command in the root
directory of your project. This starts the local web server (if it
isn\'t already running) and opens Word with your add-in loaded.

\`\`\`npm start

Webpack launches:

![](media/image12.png){width="6.5in" height="3.451388888888889in"}

\*\*\*key and cert

![](media/image13.png){width="6.5in" height="3.4208333333333334in"}

1.  In windows open power shell as "administrator"

2.  Navigate to your directory

    a.  Cd \\udex-office-add-in\\udex-author-assist

3.  Create a document in sharepoint and get the https:// url. It needs
    to be an 'online' document

4.  **share the document**

5.  **Get the url:\
    **[udex-hello-world.docx](https://sap.sharepoint.com/:w:/t/UDEx/EZA0S2JvKW9NimWzrOfs88kBMjrofv0XrpetvfmvCkbnwQ?e=lUiYa5)

<https://sap.sharepoint.com/:w:/t/UDEx/EZA0S2JvKW9NimWzrOfs88kBMjrofv0XrpetvfmvCkbnwQ?e=lUiYa5>

6.  \`\`\`npm run start:web \-- \--document {url}\
    https://sap.sharepoint.com/:w:/t/UDEx/EZA0S2JvKW9NimWzrOfs88kBMjrofv0XrpetvfmvCkbnwQ?e=lUiYa5

![](media/image14.png){width="6.5in" height="2.807638888888889in"}

7.  Click yes to allow
    [https://localhost:{port}](https://localhost:%7bport%7d)

![](media/image15.png){width="6.5in" height="2.647222222222222in"}

8.  Go to the top ribbon navigation and click the three dots

9.  On the expanded navigation \> click "show Taskpane"

> ![](media/image16.png){width="3.54122375328084in"
> height="5.655543525809274in"}

10. Click "insert Paragraph"

Scratch:

![](media/image17.png){width="6.5in" height="0.5506944444444445in"}
