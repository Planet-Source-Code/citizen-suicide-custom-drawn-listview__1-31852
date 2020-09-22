{for easier reading of this file keep word wrap turned off}
{Update 1}
IntroDuction:
	   I spent many hours trying to find a way of producing the effect that this code
	produces. I spent some time looking at using bitmaps in custom DC's, but these were
	a bit difficult to work with and far to slow. Generating the lines at run time in a
	picture box and then setting the listviews .picture property were likewise very 
	slow. So I tried to find alternatives, posting on the PSC's message board and such.
	It wasn't until i found out that you could produce custom drawn or owner drawn
	controls in VB that i produced this code.
	   This is the first of two submissions concerned with changing the appearance of
	the Listview. The second being a class (hopefully - for easy code reusabillity) 
	that impliments owner drawn listviews (and maybe other controls as well)

OverView:
 	   Windows allows the programmer to mess with its controls in all sorts of ways, 
	know by C++ coders and the like for quite a while. Most of these style changes are 
	Done at run time just before the control is created, which is why VB makes it 
	difficult to get at them (since VB creates our controls for us).
	   You can get infinite customisabillity from a control through making that control
	Owner Drawn. What does this mean? Well pretty much what it says on the tin, instead 
	of the control itself handling its drawing or redrawing, the owner of the control
	(The Form - And therefore you) is responcible for it. 
	   However, there are problems with owner drawn controls. The most obvious being 
	that you have to write the draw procedure yourself, and this procedure has to be
	very tight, as it *WILL* be called every single time the control repaints. The 
	second is that (i have heard) owner drawn controls and VB are difficult to get 
	working correctly, since VB can sometimes mess up the z-order and focus of owner
	drawn controls. This can be worked around, and Owner drawn controls are the most
	flexible and easily changed, but the coding involved is somewhat prohibitive.
	   So what is the alternative? Well Custom Drawn controls. Custom Draw is a process
	where by we tell the control how to draw certain parts of itself, and let the 
	control do all the work. VB likes it (it thinks the control is handling everything
	as it should) and we like it because the code is easier.

The Theory:
	   How do we accomplish Custom Draw? Well in this method, (partially based on the
	one outlined by Matt Hart, and another implimented by Bryan Stafford) we let VB 
	create the control as ussual, then subclass the owner form so that we can get its 
	messages. Then when the control sends a message telling us its painting it-self, we
	change those aspects that need changing.

The Practice:
	   What happens is that every time a control does something with it self it sends a
	message to its owner (the form). This message is the WM_NOTIFY message, and is 
	accompanied by the actual action that is taking place. By subclassing the controls
	owner we get these messages, and we can then decide which ones to handle ourselves
	and which to send back to VB.
	   In this case the ones we want to handle are the NM_CUSTOMDRAW messages. Again 
	along with these messages comes yet more - describing what the control has done or 
	is about to do. The messages we are concerned with are: Control PrePaint, item 
	PrePaint and item Postpaint.

The Code:
	   Lets see how the two different ways of customising the listview given in the
	sample work, firstly adding a custom highlight colour.
	Sounds as though it should be easy, but it isn't. This method works like this:
	 * Subclass the owner of the control.
	 * Just before the control paints - tell it we want to have first say in item paint
	 * on item paint, check if the item is selected (and therfore highlighted).
	 * If it is turn off the row highlight (set it as unselected).
	 * Set the foreground and background colours to our custom ones.
	 * Tell the control that we want to get messages when the item painting has 
	   finished.
	 * On Post paint, if the item was selected, re-select it.
	Since the selection was turned off during the paint procedure, but our custom
	colours were turned on, we now have a custom highlight colour, and since when the 
	item had finished painting we turned the highlight back on, the control can be 
	interacted with just as if we had done nothing at all.
	   Now how is the alternating colour scheme achieved?
	 * Subclass the owner of the control.
	 * Just before the control paints - tell it we want to have first say in item paint
	 * On Item Paint check to see if the items index is easily divisible by 2 (leaves 
	   remainder 0).
	 * If it is set the items background and foreground colours to whatever we want.
	Now on every other row (one where the number leaves remainder 0 after division, 
	i.e. 2-4-6-8) there is the custom colours set by SetCustomColour. On the other rows
	we have the default background colour (i.e. 1-3-5-7 etc). To swap the colour
	assignments over (make even rows default colours, odd rows custom colours) is 
	change this line in windowProc:
		

How to use it:
	   I have implimented these routines in a module, all you need do is use the 
	routines, or write your own based on this code.
	Add this code in the Form_Load event:
		Attach Me.hWnd, ListView
	And don't forget to UnAttach in your form_unload event:
		UnAttach me.hWnd
	
	To set whether a custom highlight colour is used call:
		UseCustomHighLight True
		or 
		UseCustomHighLight false
	To Set whether a listviews items will have alternating colours call:
		UseAlternatingColour True
		or
		UseAlternatingColour False
	To set the custom highlight colour use this code:
		Dim NewColours as ItemColourType
		NewColours.ForeGround = Your Colour Code Here
		NewColours.BackGround = Your Colour Code Here
		SetHighLightColour NewColours
	To Set the custom alternating colour (the first colour is the default foreground/
	background colours) call:
		Dim NewColours as ItemColourType
		NewColours.ForeGround = Your Colour Code Here
		NewColours.BackGround = Your Colour Code Here
		SetCustomColour NewColours

	  That should be all you need do. There are other functions and subs you can call 
	but they are relativly easy to understand.
	  Just remember that this uses subclassing, and therfore *WILL* crash your IDE
	if you break, or use the stop button. Also if you use the End statement, as it does
	no shutting down of you app or cleaning up, you should unload all forms and controls
	first. Using the End statement - at worse it'll crash your computer - at best it 
	will cause resource leaks.

	   Of course there is a lot more you can do with this code, and i have tried to make
	the module as easy to understand as possible so that future modifications are as 
	easy to accomplish as possible.

Updates:
	  There is a lot more you can do with this code. If, for instance, you were to
	store custom colours for each item in the listview you could have a custom
	background colour for every item, or a different highlight colour depending on an
	items properties.
	   Please send any updates or bug fixes to citizensuicide@aol.com, and I will
	re-upload the code with your name in the credits.

	   Fixed the bug that caused improper function with multiselect on.

DrawBacks and Bugs:
	   With multi selection on and more than one item highlighted there is an intrusive
	refresh flicker, caused by continual refresh.
	   You'll have to be careful when running in the IDE, as subclassing can crash your
	computer.
	   Obviously this is slower than letting the listview look after itself.
	   There can be problems getting the colours to display right during scrolling, they
	can appear missaligned - i know not why.
	   Highlight colour tends to dissapear on coloumns that are being resized, and will
	not reappear until the item is clicked, even though the item is still flagged as
	selected. Again i'm not sure what causes this.
	   The selection is always hidden when the listView doesn't have the focus, no 
	matter what value is set in hideselection.
	   Any other listviews on the same form will not display correctly.
	   Setting alternating colours will overlay any background picture.

--------------------------------------------------------------------------------------------
Developemnt platform:
	AMD Athlon 1333 Mhz
	640 Mbytes of RAM
	1024x768 resolution at 32bit colour [my monitor lets me down :(]
	Windows 2000 professional and Windows XP proffessional - Dual boot.
	Visual Basic 6 Service Pack 5

	No extra lag was experienced on this system.

--------------------------------------------------------------------------------------------
What You need:
	Visual Basic 5 or 6.
	
	You'll also need a copy of Internet Explore 4.7 or greater to run the sample.
--------------------------------------------------------------------------------------------
Credits:
	* Bryan Stafford of New Vision Software® - For some API declares and implimentation
	* Matt Hart - For his article in the July 1999 VBPJ > Modify Any Objects Style
	* And anyone i forgot

Sites:
	* Www.planetsourcecode.com - Of Course
	* www.vbaccelerator.com
	* www.CodeGuru.com
	* www.vbpj.com