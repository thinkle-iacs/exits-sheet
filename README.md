# Exits Sheet Automation

This Google Apps Script project lives here:

https://github.com/thinkle-iacs/exits-sheet

It manages the `Account Suspension Progress` sheet for staff exits. Form rows
provide the departing account, submitter email, expected last day, and the
number of days after exit before suspension. The script calculates warning and
suspension dates, sends notices, removes direct group memberships, suspends the
account, and flags accounts that are later re-enabled.

Managed with clasp.
