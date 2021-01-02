# README

[Google Spreadsheet Voting](http://github.com/JettJones/google_script_voting)

Author: Jett Jones

Last docs update: 2021-01-02

Previously [Instant Runoff Voting by Chris Cartland](https://github.com/cartland/instant-runoff) and
[Google Spreadsheet Voting by Darell Ross](https://github.com/eukota/google_script_voting)


# What is this Voting Script?

This script allows running a ranked choice vote from a Google Docs Form via a Google Docs Spreadsheet. All configuration is done within the Spreadsheet with little to no programming experience necessary to configure and administrate your vote.


## Ranked Choice Voting
- Wikipedia describes RCV well enough: http://en.wikipedia.org/wiki/Ranked-choice_voting

- CGPGrey's five minute video, "The Alternative Vote", also does a great job on YouTube: https://www.youtube.com/watch?v=3Y3jE3B8HsE

In this project, RCV is a method of electing one winner. Voters rank candidates in a Google Form and the administrator runs a script with Google Apps Script to determine the winner.

### Ranked choice from a voter's perspective

1. You get one vote that counts. It comes from your top choice that is still eligible.
2. If a candidate gets a majority of votes, then that candidate wins.
3. If no candidate has majority of all votes, then the candidate with the least votes is removed.
4. If your top choice is removed, the next eligible candidate on your list gets a vote. The process repeats until there is a winner.

_Notes about algorithm_

* Majority means more than half of all votes. Example: candidate A gets 3 votes and candidates B, C, and D each get 1 vote. Candidate A does not have majority because 3 is not more than half of 6.
* If multiple candidates tie for least votes, first attempt to tiebreak looking at second preferences, then all still tied candidates are removed.
* It is possible that multiple candidates tie for first place, in which case the vote ends in a tie.


# Setup a ballot
The workflow will be:
* setup a google form for the ballot
* connect it to a spreadsheet
* attach this script to the spreadsheet

## Setting Up the Ballot Form
This script assumes the structure of votes coming out of the google form.  This section covers how to setup a form to match.

1. Create a new, blank form on [forms.google.com](https://docs.google.com/forms/u/0/)
2. Add an entry for the Secret key, it should use the type: `Short answer`.
3. Add a new question for each vote, they'll use the type: `Multiple choice grid`
   * Add each candidate as a Row.
   * For the columns use `1st` `2nd` `3rd` ... etc, up to the number of candidates.
   * In the bottom right of the question there's a `...` menu with two important options to enable:
     * Limit to one response per column
     * Shuffle row order
   * Leave 'Require a response in each row` toggled off.
   * Finally, remember the name you give this question - you'll use it in spreadsheet configuration next.

## Connect to a spreadsheet

1. In the Responses tab of the form, in the (`...`) menu, `Select response destination`
2. Create a new spreadsheet.

## Attach this script

1. Open the newly created spreadsheet.
2. From the Tools menu, select "Script Editor..."
3. In the window that appears, select "Blank Script"
4. Paste the contents of instant-runoff.gs into the script.
5. Save your project with a name of your choice.
6. Return to your Google Sheet and refresh the page.


After a moment, a new "VOTING" menu should appear and the script should create several tabs
( **Configure, Results, and FilteredVotes** ).  The menu option in the VOTING menu titled "Initialize
Spreadsheet" can re-run this setup at any time.

The remaining steps will happen in that `Configure` tab - the other two are used while tallying.

## Configure the spreadsheet

There are three main columns in `Configure`: **Keys, Votes, Choice Counts**

* Keys - keys are optional, but prevent double voting or allowing more folks than you intend to use your form.
* Votes - the name, in the original form, of the question that represents a vote.
* Choice Counts - the number of candidates in the corresponding vote.

For example, if the question was titled 'Mayor' and you created three entries in the Ballot, then
you would enter 3 in the second column next to Mayor in the first column.

The use of keys can be enabled by entering `Yes` in column B1 of the configure sheet, keys are otherwise disabled.

# Running an Election

# Setting Secret Voting Keys

Giving each vote a short key ahead of time ensures:
  * each voter has only one vote entry
  * they can return and change their vote if they like.


1. Open the Elections Spreadsheet.
2. Select the “Configure” sheet.
3. On the first column under the heading of “Keys”, enter one secret key per line.
   - These keys should be distributed, one per person, to the voters. Voters enter these keys into the form.
   - Secret keys can be any non-blank value
4. Set the `B1` cell of the Configuration sheet to `Yes`.
   - this tells the tally step to only count votes with the given keys.

## Running the Vote
After all voters have their individual secret keys, distribute a link to the live form and wait for them to fill in the form.

## Tallying the Vote

From the menu “VOTING”, select the “Tally Votes” option. For less than 100 votes, this should take a
few seconds. You can watch the highlighting as it happens in the Votes sheet. When it is complete,
the Results sheet will contain the results for each vote, with details for of the rounds of counting
and elimination. A final list will also appear in a message box.


# Links

* The google form is inspired by [xFanatical](https://xfanatical.com/blog/how-to-create-ranked-choices-in-google-forms/)
* [Wikipedia Ranked Choice Voting]( http://en.wikipedia.org/wiki/Ranked-choice_voting)
* [CGPGrey "The Alternative Vote"](https://www.youtube.com/watch?v=3Y3jE3B8HsE)
