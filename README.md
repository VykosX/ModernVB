<h1>ModernVB</h1>
<p align="center">
   <img src="https://i.imgur.com/XrGDwxQ.jpg">
</p>

---

Are you still in love with Visual Basic 6 even decades after Microsoft officially abandoned all
support for the platform? Do you find yourself coding regularly in the VB6 IDE yet wishing it
didn't look so dated? Ever hoped it could look nice and pretty like modern Visual Studio editions?

Well look no further friend! Thanks to ModernVB you now can! ModernVB is a suite of unofficial
modifications for the Microsoft Visual Basic 6.0 Integrated Development Environment which completely
revamp the user interface with new icons taken from Microsoft Visual Studio (which are free to use),
new custom toolbars and many addins that unlock a great deal of extra functionality to Visual Basic.

All for the low-low price of absolutely nothing at all! :D

To stay up to date on the progress of this project, feel free to follow the discussion thread on the VBForums, where you can additionally find many more screenshots: http://www.vbforums.com/showthread.php?885405-Modernizing-the-VB6-IDE

---

**The suite is currently comprised of 5 separate elements:**

1. A Registry file for you to merge that will add modern VS icons to all the VB menus that can be themed.
    Icons were painstakingly chosen, imported and edited to best represent the functionalities of each entry.

2. A custom add-in which replaces the standard VB toolbars with replacement toolbars with the same items but
    using modern icons instead. This was necessary in order to avoid icons disappearing for items that can't
	be selected, are disabled or unavailable in the current IDE context, as the native VB toolbars are not
	capable of rendering a disabled picture for icons with more than 16 colors. 
	
	Furthermore, the addin also extends the VB6 IDE with several quality of life features to improve usability.
	
	The source code for the addin is also available and included for your enjoyment.
	
3. Patch files for VB6.exe, VBA6.dll and VBIDE.dll files. Here is the full list of changes enabled by these patches:

	- Replacing native 4-bit resources with high quality 24-bit icons and bitmaps.	
	- Minor adjustments to VB6.exe's assembly code required to force the IDE to render bitmaps in full 24-bit color depth instead of 4-bit by default.	
	- An integrated manifest file to enable visual styles for controls within the Visual Basic IDE.	
	- Changes to VBA6.dll to enable custom theme colors instead of the limited 16 color choice that VB6 supports by default	
	- Finally, many changes to VB's internal dialogs were made to enhance their usability in higher screen resolutions.
	
	Instructions to use the patches:
	
	Apply the included .xdelta patches with Delta Patcher Lite (https://github.com/marco-calautti/DeltaPatcher)
    or any other compatible xdelta3 patcher (included in the download).
	
	Distributing delta patch files avoids copyright infringement issues as none of the original copyrighted code
	is included in the patches, only a dozen or so handwritten assembly lines and official resource files which
	can be freely obtained from this address: https://www.microsoft.com/en-us/download/details.aspx?id=35825.

	In order for the patches to work correctly, ensure you're using the final SP6 for VB6 and that your local versions of the VB6 files match the following:
	
	- VB6.exe    (v6.0.97.82  - MD5: 5AC021164304F5C90C0820199D4C3E7E)
	- VBA6.DLL   (v6.0.0.9782 - MD5: CAC38827BCD9F710EAE33F949F96EF96)
	- VB6IDE.dll (v6.0.92.82  - MD5: 278E2CB9140BEEB92C5397B4D8B86D3A)
	
4. Registry files containing custom themes and settings for the VB6 IDE that can enhance the look of the software
    and provide optimal configurations for your coding experience, as well as links to a few programming oriented
	alternative monospace font choices.
	
5. Finally, a list of recommended add-ins to extend the functionality of the IDE is provided with shortcuts to
	download each one from official sources. Most of the recommended addins are free to use, although a few
	 recommendations are made for exceptional commercial addins that can greatly benefit the user experience.



**Well, that's it for now!**

I hope this pack of modifications is useful for other coders and ultimately contributes to revitalizing our love of
this fantastic, yet sadly forgotten language! Honestly, the VB6 IDE is both extremely robust and well designed.
If only someone more capable can go on to write a compatible 64-bit multiplatform compiler for VB6, then we'd well
and truly have achieved a complete ressurection of the language, with both a modern IDE interface and support for
current operating systems and the latest technologies.

For now, we can only dream and hope our efforts might one day inspire such a glorious ressurgence. :)
	
---

THIS SOFTWARE AND ALL ACCOMPANYING MATERIALS ARE PROVIDED "AS IS", IN THE HOPES THAT IT PROVES TO BE USEFUL,
WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, OR OTHERWISE, ARISING FROM, OUT OF OR
IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE AND ACCOMPANYING MATERIALS.
