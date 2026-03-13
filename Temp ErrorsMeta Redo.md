Commit current working version after changing 
class initialize to code = zero



Refactor .RecordErr, xx and y to simplify.   move to use of object oriented helper structure for lookups of attributes from Errors_ sheet table 

processing fatal errors (current ErrorHandling.RecordErr that is called in ErrorExit block of all project Boolean functions), processing non-fatal warning messages in processing calls to add informative comments to cells. 

The three used cases mentioned in the purpose, all involve reading a set of values from a found row in the projects Errors_ sheet table.  The current approach uses a sentinel value to flag, the case where the row is not found. See [[ErrorHandling Class]] for background. Current code for fatal errors is essentially a step-by-step procedure, but it follows a meandering path through RecordErr which calls AppendErrMsg which calls other sub functions to look up a Locn base row in Errors_ and compute the ErrorHandling.iCodeReport code used to then look up a specific row’s attributes. The goal of this refactoring is to streamline RecordErr into a straightforward procedure that  branches appropriately for errors that are either user facing or developer facing as flagged in the Errors_ table

* Move Errors_ lookups to a new ErrorsMeta helper class that looks up  ErrorHandling.Locn Base row and computes ErrorHandling.iCodeReport
* Detect not found Base row and malformed row data
* Use the helper class to manage lookups for fatal error reporting from .RecordErr, warning message reporting by ReportWarningMsg and comment from xxx
* Use Boolean flag attributes to seed error message generation for the three use cases (Messages could be consistently generated in the helper class and passed to ErrorHandling attributes for reporting)
* Refactor to eliminate the not found xxx sentinel value in data
* Within the helper class convert to using a tblRowsCols instance tblE for the Errors_ table. Instead of setting local column in row ranges for the table as currently, provision, the tblE instance with xxx = True to automatically set colrng’s (see xxx tblRowsCols method that does this for the errors sheet). Provision also automatically sets .rngRows attribute needed for lookups. 
* Note that RecordErr for fatal errors needs lookup to determine whether user or developer facing but other use cases are inherently user facing and mis-entered Errors row data for these codes should be automatically corrected to UserFacing = True if read in as False initially

In procPlan Flow section, show detailed pseudo code flow for the three use cases separately 
E.g. set ErrorHandling.Locn from Locn arg
Instance ErrorsMeta as meta etc.

- If not driver set CallingFunction False
- Append to .ErrMsg for pre-existing, developer facing message (case where .NewErr False and Not .IsUserFacing when RecordErr called)
- If .NewErr branch
	- Instance helper class meta
	- GetBaseErrCode (move to ErrorsMeta; flag IsBaseNotFound error)
	- If base found, Compute ErrorHandling.iCodeReport
	- meta.Load iCodeReport rows attributes by lookup set IsNotFound  
	- meta.validate to set IsMalformed error flag
	- Set local MsgNew based on either error flags or msg lookup or Not userfacing
	- Append local MsgNew to ErrMsg
	- .isNewErr = False
	- 

Instructions about testing
Chat attributes and helper class methods

Background doc updates 
Class direct attr setting. Not get/let
Add expectation of listing refs at bottom of background section
Review chat

List of comments about current functions