1
Description
My project aims to automate the token assignment process for appointment
bookings in various industries, including hospitality and passport services.
By utilizing a Google Form to collect user information and a Google Sheets
spreadsheet as a centralized database, the project assigns tokens to individu-
als based on the order of their form submission. This eliminates long waiting
times and provides users with a more efficient and convenient experience.
Once the form submission period ends, the project parses the spreadsheet
data and assigns tokens to individuals, considering factors such as individual
preferences for the time slot. Personalized emails are then generated and sent
to each individual, addressing them by name and including their assigned to-
ken number and appointment time.
Additionally, the project includes a cancellation feature that allows users
to cancel their appointments if needed. When a cancellation request is re-
ceived, the project updates the spreadsheet accordingly, freeing up the as-
signed token for reassignment. This ensures optimal utilization of available
appointment slots and allows other individuals to be accommodated in a
timely manner.
2
Implementation
Once the form is filled by the users all the data is stored in a spreadsheet
which is connected to a app script where all the implementation is done.
The form collects user preferences for the given time slots
Initially, all these responses are stored in a sheet named ”Responses” which
then is split into two sheets named ”copy” and ”cancellations” using the
query function of spreadsheets checking whether the user filled the form for
appointment booking or cancellation. The data from the sheet ”Copy” is
then modified using the removeDuplicatesAndStore() function which stores
only the first response of the form filler if they submit the form multiple
times for appointment booking and all the duplicate responses are stored
into a sheet named ”Duplicates”.
2First, all the data is parsed and at each location, a maximum of 6 peo-
ple get an appointment based on their preferences for time slots. If more
than 6 people opt for a particular location, the remaining people are sent
emails regarding the unavailability of slots and are added to a waiting list.
Later the people who opted for cancellation will get emails regarding the
cancellation of their request and then if there are people in the waiting list
in that location they get this slot.
For more information regarding query function refer[1] and for information
regarding sending emails refer[2] and [3]
3
Customization
My project is customized to
• To take the preferences of time slots from the user and give appoint-
ments accordingly
• People can opt for location preference
• addressing user with Mr and Ms
• People can fill out the form for appointment booking or cancellation
• People who could not get a slot are added to a waiting list and given
appointments based on the cancellation requests of other people at that
location
• The form is customized such that
– The initial page of the form is the same but later the next page
depends on whether the form was filled for appointment booking
or cancellation.
– The details the name entered contains only characters
– the phone number contains 10 digits
– email id is in a valid format
– every time slot is given a rank
– A note for the users telling them that submitting the form multiple
times will result in the cancellation of the request
34
Code Compilation
Using =QUERY(’Responses’ !A:N,”Select * Where N=’Appointment Book-
ing’”)
the data of the people who filled for appointment booking is stored into the
spreadsheet named Copy
using =QUERY(’Responses’ !A:N,”Select * Where N=’Cancellation’”)
the data of the people who filled for cancellation is stored in sheet named
Cancellation
Run the appointment token function in the appscript
5
Google Form and Spreadsheet
Appointment Form
Spreadsheet
References
[1] url: https://www.guidingtech.com/save-google-form-responses-
different-sheetss/.
[2] url: https://www.codementor.io/@olatundegaruba/google-apps-
script-automated-emails-m2m0ojq9v.
[3] url: https://spreadsheet.dev/send-an-email-for-every-row-
in-a-google-sheet.
