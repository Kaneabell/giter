param(
	[parameter(mandatory = $true)]
	[string]$regexpattern, 
	[string]$searchpath = $pwd,
	[string]$filetype = "*.*"
	)

function contentsearch ($pattern, $path, $type)
{
	$dir = $pwd
	cd $path
	$files = get-childitem -recurse | where { $_.name -like $type }
	foreach ($file in $files)
	{
		$content = [system.io.file]::readalltext($file.fullname)
		$match = [regex]::match($content, $pattern)
		if ($match.success)
		{
			write-host $file.fullname -foregroundcolor green
		}
	}
	cd $dir
}
#call function    
contentsearch $regexpattern $searchpath $filetype
