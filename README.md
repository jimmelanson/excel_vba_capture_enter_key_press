# excel_vba_capture_enter_key_press

You can't actually capture the ENTER key being pressed in excel VBA. You can only capture a key press
for a printable character. This gives us a hurdle to overcome if we want to use the ENTER key to execute
a procedure.

In my case, I had a multi-worksheet database and I wanted to be able to search it:

1. Without searching every worksheet seperately
2. By performing some funky concatenations within the search. (medical terms: prefix, combining forms, suffix)

I did put a button on the page to execute the search, but as this was for transcription - speed was paramount.
Having to stop and grab your mouse to click the search button slowed things down. It is much faster to press
the enter button with your pinky whilst you are typing.

My initial attempt to solve this was to assign a procedure called "SearchTerms" directly to the ENTER key like this:

<code>Application.OnKey "~", "SearchTerms"</code> Tilde is the shorthand for the keyboard enter key

<code>Application.OnKey "{ENTER}", "SearchTerms"</code> This is the numeric pad enter key

This works great as long as you remember to turn it off when the workbook is closing. The problem I have with this,
however, is that it applies to every worksheet in the workbook. This caused problems with the speed of entering
data on other worksheets. You could no longer just press return to go down to the next line.

I did try turning these <code>Application.OnKey</code> commands on and off using <code>Worksheet_Activate</code> but
it would not work. Once you turn these on, you need to turn them off at the workbook level, not the worksheet level.
This is why I came up with this simple workaround to get the ENTER key being pressed to execute the search procedure
that I had written.

On my sheet to search the medical term, I had an area for a status message and several lines for search results
to be displayed. All the cells were PROTECTED except for the cell in which you entered the search term (E2), and the
worksheet was then protected so that you could not type in anything other than the search term cell.

Change #1: I inserted a row directly beneath the row (Row 2) in which the user entered search term cell appeared.
Change #2: I unlocked one cell directly beneath the user entered search term cell (E3).

With those two changes made, every time I entered a term and then pressed the enter key, the cursor would advance
from cell E2 to E3. I then added this code to the worksheet in the VBA editor:

<code>Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Not Intersect(Target, Range("E3")) Is Nothing Then
    
        Call SearchTerms
        
    End If
    
End Sub</code>
blah blah

https://github.com/jimmelanson/excel_vba_capture_enter_key_press/blob/bd617354909250e64f3539e9bbdcd9d7128f4f7c/selection_change

It's not elegant, but it works.
