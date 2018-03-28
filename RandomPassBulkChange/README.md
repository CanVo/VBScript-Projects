
# RandomPassBulkChange

A quick script to change the passwords of all local users on a Windows system. The passwords will be randomly generated and will be set at a length according to what the user wants. The script utilizes the visual basic scripting platform to perform the action and changes.

### Prerequisites

An administrator level user to be able to run the script.

### Installing

Download the "RandomPassBulkChange.vbs" file and place to where desired. After doing so, edit the follow lines to your needs:

7:  Password length.
24: User that will not have their password changed. This user is usually the account you're on so you can log back in easily to check if
    the passwords were changed properly.
26: The file path and name of the .csv file that will be created.

End with an example of getting some data out of the system or using it for a little demo

## Deployment

If on a Windows domain, it is possible to push and deploy the script through a group policy. I haven't done this yet but this idea will be put on a todo list. Results on this experiment will be given as soon as possible.

## Built With

* [VBScript](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/vb-script/t0aew7h6(v%3dvs.84))

## Known Bugs

During the execution of the script, it sometimes results in an error:


Just run the script again until the error doesn't appear and it should execute successfully. The fix will be on the todo list.

## To-do

* Test and push script through a Windows domain.
* Fix the error that sometimes occur during runs.

## Authors

* **Can Vo**


## Acknowledgments

* Kudos to anyone who's code was used as reference.
