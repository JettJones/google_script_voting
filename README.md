# README


[Google Spreadsheet Voting](http://github.com/eukota/google_spreadsheet_voting)

Author: Darrell Ross

Last docs update: 2015-04-24

Original Script: [Instant Runoff Voting by Chris Cartland](https://github.com/cartland/instant-runoff)

# What is this Voting Script?

This script allows running a vote from a Google Docs Form via a Google Docs Spreadsheet. All configuration is done within the Spreadsheet with little to no programming experience necessary to configure and administrate your vote.

Two voting styles are provided:
* Instant Runoff Voting.
* First Passed The Post Voting.

## Instant Runoff Voting 
Wikipedia describes IRV well enough: http://en.wikipedia.org/wiki/Instant-runoff_voting
CGPGrey's five minute video, "The Alternative Vote", also does a great job on YouTube: https://www.youtube.com/watch?v=3Y3jE3B8HsE

In this project, IRV is a method of electing one winner. Voters rank candidates in a Google Form and the administrator runs a script with Google Apps Script to determine the winner.

### Instant-runoff voting from a voter's perspective

1. You get one vote that counts. It comes from your top choice that is still eligible.
2. If a candidate gets a majority of votes, then that candidate wins.
3. If no candidate has majority of all votes, then the candidate with the least votes is removed.
4. If your top choice is removed, the next eligible candidate on your list gets a vote. The process repeats until there is a winner.

_Notes about algorithm_

* Majority means more than half of all votes. Example: candidate A gets 3 votes and candidates B, C, and D each get 1 vote. Candidate A does not have majority because 3 is not more than half of 6.
* If multiple candidates tie for least votes, then all are removed.
* It is possible that multiple candidates tie for first place, in which case the vote ends in a tie.

### First Passed The Post Voting

Wikipedia describes FPTP well enough: http://en.wikipedia.org/wiki/First-past-the-post_voting
CGPGrey's five minute video, "The Problems with First Passed the Post Voting", also does a great job on YouTube: https://www.youtube.com/watch?v=s7tWHJfhiyo

# Setting up the Script
To get the script configured for your spreadsheet, follow these steps:

1. Make a new Google Sheet.
2. From the Tools menu, select "Script Editor..."
3. In the window that appears, select "Blank Script"
4. Paste the contents of instant-runoff.gs into the script.
5. Save your project with a name of your choice.
6. Return to your Google Sheet and refresh the page.

After a moment, a new "VOTING" menu should appear and it should also have automatically run the first menu option in the VOTING menu titled "Initialize Spreadsheet". 

# Running an Election
The Election Spreadsheet is designed to handle two forms of voting:

* First Past The Post (FPTP)
  FPTP voting is the most commonly used voting. It suffers from various issues, most notably, the spoiler effect which requires people to vote strategically against those they don’t want to win.
* Instant-Runoff Voting (IRV)
  IRV is useful to ensure a more fair result since it eliminates the spoiler effect and allows people to freely vote for who they want without the need to strategically vote for the lesser of two evils.

 
## Setting Up the Ballot Form

The major difference between FPTP and IRV on the Ballot is that FPTP needs only a single multiple choice item for each vote type while IRV needs to have the same number of multiple choice items as there are options for that vote type.

Both voting types begin by clearing out the ballot:

1. Open the Elections Ballot form.
2. Delete all options for each vote.

### FPTP Ballot Form Setup

1. The process from here on is the same for each vote.
	a. Add a field for the first vote.
		i. Add a “Multiple Choice” item.
		ii. Name the item based on the vote (eg: “President”).
		iii. Set the Question Type to “Choose from a list”
		iv. Enter in each candidate’s name as a single entry to the list.
		v. Check the “Required question” checkbox.

### IRV Ballot Form Setup

1. The process from here on is the same for each vote:
  * Add the first option for the first vote:
    * Add a “Multiple Choice” item.
	* Name it “1st choice”.
	* Set the Question Type to “Choose from a list”
	* Enter in each candidate’s name as a single entry to the list.
	* Check the “Required question” checkbox.
	* Expand the Advanced Settings area and check the “Shuffle option order” checkbox.
  * To create the next position, click the duplicate button (it looks like two pieces of paper on top of one another in the upper right of the field).
  * Modify the duplicate so that it is not required and name it “2nd choice”.
  * Further choices can be easily duplicated from the 2nd choice and renamed to “3rd choice”, “4th choice”, “5th choice”, etc.

## Setting Up the Spreadsheet
There is only one small difference between setting up the Elections Spreadsheet for FPTP versus IRV and that is the Choice Count column in the Configure sheet.

Initial setup is the same for both voting systems. 
1. Open the Elections Spreadsheet.
2. Delete the Configure sheet.
3. If you have changed the entries on the form, then the voting will fail to take notice of them. You must reset the form’s attachment to the spreadsheet so that the entries are all in order. To do so, do the following:
	a. From the menu “Form”, select “Unlink Form”. Confirm that you are ok with this.
	b. Delete the Votes sheet. This will clear out all previous configurations of the Ballot. The Results sheet should be the only one remaining. If you lack a sheet to leave remaining, create a temporary empty one so you can delete the Votes sheet.
	c. Open the Elections Ballot form.
	d. Click the “View Responses” button.
	e. In the dialog that appears, select “New sheet in an existing spreadsheet...” and click “Choose”.
	f. Choose the Elections Spreadsheet and click “Select”.
	g. Select all rows from 2 onward that have data in them and delete them (unless you want that old data).
	h. Delete the temporary sheet you created earlier in this process.
4. From the menu “VOTING”, select “Initialize Spreadsheet”. This will recreate the Configure sheet empty.
5. Populate the Configure sheet. To help you out, some of these instructions will show up on the notes for each heading that is populated.
	a. Keys - keys allow you to make sure people do not overrun your voting system. Enter one valid key per cell. If you are not using keys, then change the second row from the default of “yes” to anything else.
	b. Votes - votes are the various votes you are holding. If you are voting on only one decision, then you will enter only one vote. Make sure to list the votes in the order you have them in your form.
	c. Choice Counts - must be configured differently depending on your voting system.
		i. FPTP - Enter a 1 for each item. This is because the voter is allowed only a single vote.
		ii. IRV - Enter the number of choices which you have configured for each vote in your form. For example, if there were three choices for President and you created the three entries in the Ballot, then you would enter a 3 in the second column next to the President entry from the first column.
6. From the menu “VOTING”, select “Setup Voting”. This will create a new sheet named Results which is also where the results will be stored.

## Setting Secret Voting Keys
Secret keys are used the same in both FPTP and IRV. 

1. Open the Elections Spreadsheet.
2. Select the “Configure” sheet. It should be the second one.
3. On the first column under the heading of “Keys”, enter one secret key per line. These keys should be distributed, one per person, to the voters. Voters enter these keys into the form. This ensures that each voter:
	a. has only one vote entry
	b. can return and change their vote as often as they like
4. Secret keys must be exactly 5 characters long and may contain lowercase or uppercase letters or any number.

## Running the Vote
After all voters have their individual secret keys, distribute a link to the live form and wait for them to fill in the form.

## Tallying the Vote
Vote tallying depends on the vote style you chose. The vote style is something you will have configured earlier. 

From the menu “VOTING”, select the “Tally Votes” option. This make take a few minutes. You can watch the highlighting as it happens in the Votes sheet. When it is complete, the Results sheet will contain the results for each vote as a note on the top column the vote is associated with. A final list will also appear in a message box.

