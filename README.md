# Excel-VBA-Sample-WithEvents-with-Class
UserDefined Event-handlers for Multiple Controls of Same Type by Using Class(es) employing WithEvents
On April 16th, 2021, someone on Reddit asked a question about how to write efficient code for multiple controls mousedown and mouse up handlers.</br>
I have decided to help answer that question.
But after I posted my answer, that post was deleted.</br>

Therefore, in order not to waste my time and energy thinking about a solution for that question, I am hereby sharing that code.</br>

The .frm and .frx files should be place in the same folder. But only .cls and .frm needs importing from VBIDE-File-Import menu.
![Naming_UserForm_Controls](Images/Userform_for_Class_example.png)</br>
The UserForm controls should be renamed as in the image above.

The code contains 2 methods: 
<ul>
  <li>Method1 is simpler method not using the class module.</li>
  <li>Method2 uses a class module to facilitate control event handling from one sub.</li>
</ul>
This behavior can be control from the UserForm ToggleButton.</br>

An extra barebone version is also included but commented and if the barebone version is required, the non-barebone version should be commented out.
The commenting and un-commenting <b>must</b> be done in both UserForm code module and Class module.
