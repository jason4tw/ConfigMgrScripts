Set env = CreateObject("Microsoft.SMS.TSEnvironment")

Wscript.echo "Dump of All Task Sequence Variables"
For Each TSvar In env.GetVariables
	Wscript.echo "Variable " & TSvar & " = " & env(TSvar)
Next
	
