# 1. Creating Export for Mail Rules
- Creating Outlook VBA script to export Outlook rules after selecting mailbox within Outlook
- Not all rules/conditions/actions are covered yet.

## Some notes:

As of Microsoft® Outlook® for Microsoft 365 MSO (Version 2508)

Available condition types for `OlRuleConditionType`
`RuleCondition.ConditionType` property returns one of these constants to identify the type of condition in a rule's `Conditions` collection.
Cast the generic `RuleCondition` object to its specific corresponding object (e.g. a condition with text, like `olConditionSubject`, 
is cast to a `TextRuleCondition` object) to retrieve the actual values, such as the words in the subject or the specific recipient name.

Constant|Description
|-|-
olConditionAccount|Message is received through the specified account.
olConditionBody|Body contains specific words.
olConditionBodyOrSubject|Body or subject contains specific words.
olConditionCategory|Message is assigned to the specified category.
olConditionCc|Message has your name in the Cc box.
olConditionFrom|Sender is in the recipient list specified.
olConditionFromAnyRssFeed|Message is generated from any RSS subscription.
olConditionFromRssFeed|Message is generated from a specific RSS subscription.
olConditionHasAttachment|Message has one or more attachments.
olConditionImportance|Message is marked with the specified level of importance.
olConditionLocalMachineOnly|Rule can run only on the local machine.
olConditionMessageHeader|Message header contains specific words.
olConditionOnlyToMe|Message is sent only to you.
olConditionRecipientAddress|Recipient address contains specific words.
olConditionSenderAddress|Sender address contains specific words.
olConditionSenderInAddressBook|Sender is in the address book specified.
olConditionSentTo|"Sent to recipients (To| Cc) are in the recipient list specified."
olConditionSubject|Subject contains specific words.
olConditionTo|Your name is in the To box.
olConditionToOrCc|Message has your name in the To or Cc box.

Available Rules for `OlRuleActionType`
|Constant|Value|Action Description
|-|-|-
olRuleActionMoveToFolder|1|Moves the message to the specified folder.
olRuleActionAssignToCategory|2|Assigns the message to one or more specified categories.
olRuleActionDelete|3|Deletes the message (moves it to Deleted Items).
olRuleActionDeletePermanently|4|Permanently deletes the message.
olRuleActionCopyToFolder|5|Copies the message to the specified folder.
olRuleActionForward|6|Forwards the message to the specified recipients.
olRuleActionForwardAsAttachment|7|Forwards the message as an attachment to the specified recipients.
olRuleActionRedirect|8|Redirects the message to the specified recipients.
olRuleActionServerReply|9|Has the server reply using a specified mail item or template.
olRuleActionTemplate|10|Replies using the specified template (.oft) file.
olRuleActionFlagForActionInDays|11|Flags the message for follow-up action in the specified number of days.
olRuleActionFlagColor|12|Flags the message with a specified colored flag.
olRuleActionFlagClear|13|Clears the message flag.
olRuleActionImportance|14|"Marks the message with the specified Importance (High| Normal| or Low)."
olRuleActionSensitivity|15|Marks the message with the specified level of sensitivity.
olRuleActionPrint|16|Prints the message to the default printer.
olRuleActionPlaySound|17|Plays a specified .wav sound file.
olRuleActionStartApplication|18|Starts a specified executable (.exe) file.
olRuleActionMarkRead|19|Marks the message as read.
olRuleActionRunScript|20|Starts a script (deprecated in modern Outlook versions).
olRuleActionStop|21|Stops processing any further rules for this message.
olRuleActionCustomAction|22|Performs a custom action (usually related to specific third-party add-ins).
olRuleActionNewItemAlert|23|Displays a custom text in the New Item Alert dialog box.
olRuleActionDesktopAlert|24|Displays a desktop alert.
olRuleActionNotifyRead|25|Requests a read notification for the message being sent.
olRuleActionNotifyDelivery|26|Requests a delivery notification for the message being sent.
olRuleActionCcMessage|27|Ccs the message to specified recipients (for send rules).
olRuleActionDefer|28|Defers the delivery of the message by a specified number of minutes (for send rules).
olRuleActionClearCategories|30|Clears all categories assigned to the message.
olRuleActionMarkAsTask|41|Marks the message as a task for follow-up.
olRuleActionUnknown|0|Unrecognized rule action (used for error handling).

Other features
- For troubleshooting purposes, choice between Immediate Window output and Excel
- Excel output with some rudimentary formatting

Current Prerequisites
- local Microsoft® Outlook® for Microsoft 365 MSO installation on machine (testing for windows only), might work for other/earlier versions but no guarantees.
- Enable Developer Mode
- Allow Macros/Code execution on
- Module Creation
