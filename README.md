# Word-Processor-Macros
A breakdown of all the components thus far:

1. "Specific Heading Level" is the Heading Level Specific code that the macro is based on. It inserts a run-in style separator before the first period of every paragraph of a specified heading style. This can be used changed by editing "SPECIFY HEADING STYLE". 

2. "Show Macro" allows the Macro to be called from the list when clicking the "Macros" button. 

3. "Form Options" calls and executes the macro based on the headings selected in the "StyleSeparatorForm".

4. "StyleSeparatorForm.frm" is a checkbox form that allows the user to specify which heading styles the macro will run on.

5. "AutoInsertingRunInStyleSeparator.bst" is the current and complete module. 
