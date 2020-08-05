# http://www.imsglobal.org/content/packaging/index.html
#
# I have started the work of dealing with calculated scores but this
# script does not currently do anyting with calculated scores!


#script params
param (

  [Parameter(Mandatory,ValueFromPipeline)]
  #[ValidateScript({ ($_ -match '\.zip$') -and (Test-Path -Path $_ -PathType Leaf); })]
  [string]$bbArchive,
  [switch]$extractDocuments,
  [switch]$renameOld = $true,
  [switch]$viewResult = $true
  
)

#these are required some of the functions
Add-Type -assembly System.IO.Compression.Filesystem;
Add-Type -AssemblyName System.web;

$Global:buffer = @();
$Global:outcome_items = [ordered]@{}; #ordered hash of columns (outcome definitions) by OUTCOMEDEFINITION.id
$Global:attempt_items = @{}; #hash of responses (attempts) by ATTEMPT.id
$Global:content_items = @{};  #hash of content objects by OUTCOMEDEFINITION.CONTENTID

#region  functions

function Push {

  param (
    [string]$item,
    [switch]$flush
  )

  $Global:buffer += $item;

  if($flush -or $Global:buffer.Length -ge 500) {
    $Global:buffer | Out-File $Global:htmlFile -Append -Encoding UTF8;
    $Global:buffer = @();
  }

}


function Push-DIV {

  param (

    [string]$class,
    [string]$name,
    [string]$value,
    [string]$toolTip

  )


  if($name) { $name = "<b>$name</b>"; }

  if($name -and -not $value) {
    $value = '<span class="empty">none</span>'; 
  }

  if($toolTip) {

    Push ('<div class="{0}" title="{3}">{1} {2}</div>'-f $class, $name, $value, $toolTip);

  } else {

    Push ('<div class="{0}">{1} {2}</div>'-f $class, $name, $value);

  }

}


function Push-Comment { 

  param (
    [string]$comment,
    [switch]$tee
  )

  Push ('<!--{0}-->' -f $comment); 

  if($tee) { $comment; };

}

function Get-EncodedHTML {

  param( 
    [Parameter(Mandatory,ValueFromPipeline)]
    [string]$text
  )

  return [System.Web.HttpUtility]::HtmlEncode($text);

}

function Get-ZipEntry {

  param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$zipPath,
    [Parameter(Mandatory)]
    [string]$entryPath
  )

  Write-Debug 'Get-ZipEntry';
  Write-Debug ('Zip:  ' + $zipPath);
  Write-Debug ('Item: ' + $entryPath);

  $zipPath = Resolve-Path $zipPath;

  $zip = [io.compression.zipfile]::OpenRead($zipPath);  

  if(-not $zip) { 
    throw ('Failed to open Zip archive: ' + $zipPath);
    return;
  }

  #convert to backslash to slash
  $entryPath = $entryPath -replace '\\', '/';

  return $zip.GetEntry($entryPath);
        
}


function Get-ZipEntryBytes {

  param (

    [Parameter(Mandatory,ValueFromPipeline,Position='0',ParameterSetName='byEntry')]
    [System.IO.Compression.ZipArchiveEntry]$zipEntry,

    [Parameter(Mandatory,ValueFromPipeline,Position='0',ParameterSetName='byPath')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$zipPath,

    [Parameter(Mandatory,Position='1',ParameterSetName='byPath')]
    [string]$entryPath

  )

  Write-Debug 'Get-ZipEntryBytes';
  Write-Debug ('Zip:  ' + $zipPath);
  Write-Debug ('Item: ' + $entryPath);

  if(-not $zipEntry) {

    $zipEntry = Get-ZipEntry $zipPath $entryPath;

  }

  $len = $zipEntry.Length;

  Write-Verbose ('Zip Entry Length: ' + $zipEntry.Length);
  
  $o = $zipEntry.Open();

  if(-not $o) {
    Write-Verbose ('Failed to get item from zip');
    return [byte[]]@();
  }

  if($o.CanRead) {

    $buff = New-Object byte[] $len;
    $i = $o.Read($buff,0,$len);

  }

  Write-Debug ('Read {0} bytes from Zip Entry' -f $i);

  $o.Close();

  return $buff
    
}


function Get-ZipEntryList {

  param (

    [Parameter(Mandatory,ValueFromPipeline)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$zipPath

  )

  $zipPath = Resolve-Path $zipPath;

  Write-Debug 'Get-ContentFromZipItem';
  Write-Debug ('Zip:  ' + $zipPath);

  $zip = [io.compression.zipfile]::OpenRead($zipPath);  

  if(-not $zip) { 
    throw ('Failed to open Zip archive: ' + $zipPath);
    return;
  }

  return $zip.Entries;
      
}


function Get-BytesAsUTF8 {

  param (
    [Parameter(Mandatory,ValueFromPipeline)]
    [byte[]]$bytes
  )

  return [System.Text.Encoding]::UTF8.GetString($bytes);

}


function Get-ZipEntryString {

  param (

    [Parameter(Mandatory,ValueFromPipeline,Position='0',ParameterSetName='byEntry')]
    [System.IO.Compression.ZipArchiveEntry]$zipEntry,

    [Parameter(Mandatory,valueFromPipeline,Position='0',ParameterSetName='byPath')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$zipPath,

    [Parameter(Mandatory,Position='1',ParameterSetName='byPath')]
    [string]$entryPath

  )

  if($zipEntry) {

    return [System.Text.Encoding]::UTF8.GetString((Get-ZipEntryBytes $zipEntry));

  }

  return [System.Text.Encoding]::UTF8.GetString((Get-ZipEntryBytes $zipPath $entryPath));

}


function Out-Binary {

  param (

    [Parameter(Mandatory)]
    [string]$FilePath,
    [Parameter(Mandatory,ValueFromPipeline=$true)]
    [byte[]]$Bytes
  
  )

  Write-Debug 'Out-Binary';
  Write-Debug ('FilePath: ' + $FilePath);
  Write-Debug ('Length: {0} bytes' -f $Bytes.Length);

  $oDir = Split-Path $FilePath -Parent;
  $oFile = Split-Path $FilePath -leaf;
  $absPath = Join-Path (Resolve-Path $oDir) $oFile;

  Write-Debug ('absPath: ' + $absPath);


  if($Bytes.Count -lt 2) { 
     
     Write-Host ('No data to write to: ' + $absPath) -ForegroundColor Yellow;
     return;

  }
  
  Write-Verbose ('Writing {0} bytes to: {1}' -f $Bytes.Count, $absPath);
  [io.file]::WriteAllBytes($absPath, $Bytes);
  
}

function Save-ZipEntryToFile {

  param (

    [Parameter(Mandatory,Position='0',ParameterSetName='byEntry')]
    [System.IO.Compression.ZipArchiveEntry]$zipEntry,

    [Parameter(Mandatory,Position='0',ParameterSetName='byPath')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$zipPath,

    [Parameter(Mandatory,ParameterSetName='byPath')]
    [string]$entryPath,

    [Parameter(Mandatory,ParameterSetName='byEntry')]
    [Parameter(Mandatory,ParameterSetName='byPath')]
    [string]$OutputPath

  )

  if(-not $entry) {
    Out-Binary -Bytes (Get-ZipEntryBytes -zipPath $zipPath -entryPath $entryPath) -FilePath $OutputPath;
    return;
  }

  Out-Binary -Bytes (Get-ZipEntryBytes -zipEntry $zipEntry) -FilePath $OutputPath;

}


function Export-GBSummary {

  $col_columns = [ordered]@{};
  $i = 0;
  $outcome_items.Values | Sort-Object position |
    ForEach-Object {
      if($_.is_calculated -match 'false') { 
        $col_columns[$_.id] = '{0}. {1} ({2})' -f ++$i, $_.title, $_.points_possible;
      }
    }

  $stu_columns = @('student','student_id')

  [string[]]$colNames = $stu_columns + $col_columns.Values;  

  $gb = @();
  foreach ($stu in ($roster.Values | Sort-Object name)) {
    if($stu.role -match 'student') {
      $o = New-Object pscustomobject | Select-Object $colNames;
      $o.student = $stu.name;
      $o.student_id = $stu.student_id;
      foreach ($key in $col_columns.Keys) {
        if($stu.grades.ContainsKey($key)) {
          $aid = $stu.grades[$key];
          $s = $attempt_items[$aid].score;
          $g = $attempt_items[$aid].grade;
          if($g) {
            $o.$($col_columns[$key]) = '{0}/{1}' -f $s, $g;
          } else {
            $o.$($col_columns[$key]) = $s;
          }
        } else { $o.$($col_columns[$key]) = '0'; } 
      }
      $gb += $o;
    }
  }

  $gb | Export-Csv $gbSummaryFile -NoTypeInformation -Encoding UTF8;

} #Export-GBSummary

#endregion functions

#$DebugPreference = 'Continue'; #debug
#$VerbosePreference = 'Continue'; #debug

$extracted = (Get-Date).ToString('yyyy-MM-dd@HH:mm:ss');

#region files and folders

$bbaFile = Get-Item $bbArchive;
$bbaName = ($bbaFile.BaseName -replace 'ArchiveFile_','');
$outputFolder = Join-Path $bbaFile.DirectoryName ('bbaExtraction_' + $bbaName);
$outputFileFolder = Join-Path $outputFolder 'files';
$Global:htmlFile = Join-Path $outputFolder ($bbaName + '.html'); #used in local functions
$Global:gbSummaryFile = Join-Path $outputFolder ($bbaName + '.csv'); #used in local functions

$localCSS = 'bbaExtractorEX.css';
$outputCSS = Join-Path $outputFolder $localCSS;

if(-not(Test-Path $localCSS)) {
  throw ('CSS file not found: ' + $localCSS);
}

@"
bbaExtractorEX.ps1 - jorgie@missouri.edu 2017/8, adapted for Ferris State University 2020 by Tails Solutions LLC

  Course Name:  $bbaName;
Input Archive:  $($bbaFile.FullName)
Output Folder:  $outputFolder;
  Output HTML:  $(Split-Path $htmlFile -Leaf)
   Output CSS:  $(Split-Path $outputCSS -Leaf)
 Files Folder:  $(Split-Path $outputFileFolder -Leaf)

"@

if(Test-Path $outputFolder) { 
  if($renameOld) {

    $oldFolder = '';
    $oldAddition = 0;
    do { $oldFolder = '{0}.({1:000})' -f $outputFolder, $oldAddition++ }
    until (-not(Test-Path $oldFolder))
    'Renaming old folder: ' + $oldFolder;
    Rename-Item $outputFolder $oldFolder -ErrorAction stop;

  } else {

    throw ('Output folder exists, please rename or delete it!');
    return;

  }
} 

if(-not(Test-Path $outputFolder)) {
  $null = New-Item -Path $outputFolder -ItemType Directory -ErrorAction Stop;
}

if(-not(Test-Path $outputCSS)) {
  Copy-Item $localCSS $outputCSS -ErrorAction Stop;
}

if($extractDocuments -and -not(Test-Path $outputFileFolder)) {
  $null = New-Item -Path $outputFileFolder -ItemType Directory -ErrorAction Stop;
}

#endregion files and folders


##bb:title values
$courseSettingsType = 'course/x-bb-coursesetting'; #single
$userResType = 'course/x-bb-user';                 #single
$cmResType = 'membership/x-bb-coursemembership';   #single
$gbResType = 'course/x-bb-gradebook';              #single
$afResType = 'course/x-bb-attemptfiles';           #multi
$dbResType = 'resource/x-bb-discussionboard';      #multi

#region get the manifest
$manEntry = Get-ZipEntry $bbaFile.Fullname 'imsmanifest.xml';
'Manifest: {0} LEN: {1} MOD: {2}' -f $manEntry.Name, $manEntry.Length, $manEntry.LastWriteTime.LocalDateTime;
[xml]$manXml = Get-ZipEntryString $manEntry;
if(-not $manXml.manifest) { throw 'Failed to get manifest'; }

#endregion get the manifest

#region course info
$csRes = $manXml.manifest.resources.resource |
  Where-Object { $_.type -eq $courseSettingsType; };
'Course Settings: ' + $csRes.file;
# Do a count of nodes. If > 1, use first as that should correctly define file we want
$csFile;
if ($csRes.HasChildNodes) {
  $csFile = $csRes.file[0];
} else {
  $csFile = $csRes.file;
}

[xml]$csXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $csFile;
if(-not $csXml.COURSE) { throw 'Failed to get course info'; }

$courseInfo = [pscustomobject][ordered]@{
  'title' = $csXml.COURSE.TITLE.value;
  'date_created' = $csXml.COURSE.DATES.CREATED.value;
  'date_updated' = $csXml.COURSE.DATES.UPDATED.value;
  'date_start' = $csXml.COURSE.DATES.COURSESTART.value;
  'date_end' = $csXml.COURSE.DATES.COURSEEND.value;
  'exported' = $manEntry.LastWriteTime.LocalDateTime;
}

Write-Verbose ('CourseInfo' + ($courseInfo | Format-Table | Out-String));

#$csXml = $null; #cleanup

#endregion course info


#region roster
$userRes = $manXml.manifest.resources.resource | Where-Object { $_.type -eq $userResType; };
[xml]$uXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $userRes.file;
if(-not $uXml.USERS) { throw 'Failed to get user info'; }

$roster = @{}; #hash of user objects by USER.id

$uXml.USERS.USER | ForEach-Object {

  $roster[$_.id] = [pscustomobject][ordered]@{
    'id' = $_.id;
    'role' = '';
    'username' = $_.USERNAME.value;
    'student_id' = $_.STUDENTID.value;
    'name' = ($_.NAMES.GIVEN.value, $_.NAMES.MIDDLE.value, $_.NAMES.FAMILY.value) -join ' ';
    'email' = $_.EMAILADDRESS.value;
    'date_created' = $_.DATES.CREATED.value;
    'date_updated' = $_.DATES.CREATED.value;
    #'portal_role' = $_.PORTALROLE.ROLEID.value;
    #'system_role' = $_.SYSTEMROLE.value;
    'grades' = @{};
  }

}

'Users: {0} {1}' -f $roster.Keys.Count, $userRes.file;

#$uXml = $null; #clean up

#add an unknown user

$dummyUserId = '_000000_0';

$roster[$dummyUserId] = [pscustomobject][ordered]@{
  'id' = '_000000_0';
  'role' = 'Unknown';
  'username' = '';
  'student_id' = '';
  'name' = 'Unknown';
  'email' = '';
  'date_created' = '';
  'date_updated' = '';
  #'portal_role' = 'Unknown';
  #'system_role' = 'Unknown';
}

#endregion roster


#region course memberships
$cmRes = $manXml.manifest.resources.resource | Where-Object { $_.type -eq $cmResType; }
[xml]$cmXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $cmRes.file;
if(-not $cmXml.COURSEMEMBERSHIPS) { throw 'Failed to load course memberships'; }

$memberships = @{}; #hash of membership objects by COURSEMEMBERSHIP.id

$cmXml.COURSEMEMBERSHIPS.COURSEMEMBERSHIP | ForEach-Object {
  
  $id = $_.id;
  $user_id = $_.USERID.Value;
  $role = $_.ROLE.Value;

  $memberships[$id] = [pscustomobject][ordered]@{
    'id' = $id;
    'user_id' = $user_id;
    'role' = $role;
  }

  if($roster.ContainsKey($user_id)) { $roster[$user_id].role = $role; }

}

'Memberships: {0} {1} ' -f $memberships.Keys.Count, $cmRes.file;

#$cmXml = $null; #cleanup

#endregion course memberships


#region attemp files
$attemptFiles = @{}; #hash of lists of file objects by ATTEMPTFILE.ATTEMPTID aka ATTEMPT.id

$bucketIndex = 0;
$bucketCount = ($manXml.manifest.resources.resource |
  Where-Object { $_.type -eq $afResType; }).Count

$p1Name = 'Scanning File Resources';

$p2Name = 'Scanning Attempt Files';
if($extractDocuments) {
  $p2Name = 'Extracting Attempt Files';
}

$manXml.manifest.resources.resource |
  Where-Object { $_.type -eq $afResType; } |
    ForEach-Object {

    $base = $_.base;
    $aFile = $_.file | Select-Object -First 1;

    Write-Progress $p1Name -Status ('{0}/{1}' -f ++$bucketIndex, $bucketCount) -PercentComplete ($bucketIndex/$bucketCount*100) -Id 1;

    [xml]$afXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $aFile;

    $afCount = ($afXml.ATTEMPTFILES.ATTEMPTFILE).Count;
    $afIndex = 0;
    $usedFilenames = @();

    $afXml.ATTEMPTFILES.ATTEMPTFILE | ForEach-Object {

      $aid = $_.ATTEMPTID.Value;

      Write-Progress $p2Name -Status ('{0}/{1}' -f ++$afIndex, $afCount) -PercentComplete ($afIndex/$afCount*100) -Id 2;

      $afo = [pscustomobject][ordered]@{
        'base' = $base;
        'id' = $_.id;
        'attempt_id' = $_.ATTEMPTID.Value;
        'file_name' = $_.FILE.NAME #[System.Web.HttpUtility]::HtmlEncode($_.FILE.NAME);
        'file_link_name' = $_.FILE.LINKNAME.value;
        'file_size' = $_.FILE.SIZE.value;
        'file_action' = $_.FILE.FILEACTION.value;
        'file_entry' = ($base,$_.id,$_.FILE.NAME) -join '/';  
        'date_created' = $_.FILE.DATES.UPDATED.Value;
        'date_updated' = $_.FILE.DATES.CREATED.Value;
        'extracted_path' = '';
      }
      
      if($attemptFiles.ContainsKey($aid)) {

        $attemptFiles[$aid] += $afo;

      } else {

        $attemptFiles[$aid] = @($afo);

      }

      $safeName = $afo.file_name -replace '[^a-z0-9_.#-]', '';

      $target = Join-Path $outputFileFolder ('{0}_{1}_{2}' -f $afo.base, $afo.id, $safeName);


      #deal with possible duplicate filename (should not happen since filename has file_id)
      $i = 0;
      if($usedFilenames.Contains($target)) {
        do { $target = Join-Path $outputFileFolder ('{0:00}_{1}_{2}' -f $i++, $afo.id, $safeName); }
        while ($usedFilenames.Contains($target))
      }
      $usedFilenames += $target;

      $relTarget = $target.Substring($outputFolder.Length + 1);

      $afo.extracted_path = $relTarget;

      if($extractDocuments) {
        Save-ZipEntryToFile -zipPath $bbaFile.FullName -entryPath $afo.file_entry -OutputPath $target;
      }

    } #ATTEMPTFILES.ATTEMPTFILE

    Write-Progress $p2Name -Id 2 -Completed;

  } #files of type $afResType

'Attempts with files: ' + $attemptFiles.Count;

Write-Progress $p1Name -Id 1 -Completed;

#endregion course memberships

#region process_gradebook

#region load gradebook
$gbRes = $manXml.manifest.resources.resource | Where-Object { $_.type -eq $gbResType; }
[xml]$gbXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $gbRes.file;
if(-not $gbXml.GRADEBOOK) { throw 'Failed to load gradebook'; }
$ocdCount = $gbXml.GRADEBOOK.OUTCOMEDEFINITIONS.OUTCOMEDEFINITION.Count;
#$ocdIndex = 0;
'Gradebook: {0} Outcomes, {1}' -f $ocdCount, $gbRes.file;
#endregion load gradebook

$fix_comments = '(</*span[^>]*>|<p[^>]*>|&nbsp;|</p>|^<br>|<br>$)';
$fix_content = "( ^ +|<[^>]*>|</div>|`n| +$)";

$p3Name = 'Scanning Gradebook';

$gbXml.GRADEBOOK.OUTCOMEDEFINITIONS.OUTCOMEDEFINITION | 
  Sort-Object @{Expression={$_.POSITION.Value}} | 
    ForEach-Object {

      if(-not $_) { return; }

      Write-Progress $p3Name -Status $_.TITLE.value -Id 3;

      $oOCD = $_;

      $ocd_id = $oOCD.id;

      #region define outcome definition, aka ocd, aka column

      $ocd = [pscustomobject][ordered]@{
        'id' = $oOCD.id;
        'category_id' = $oOCD.CATEGORYID.value;
        'scale_id' = $oOCD.SCALEID.value;
        'secondary_scaleid' = $oOCD.SECONDARY_SCALEID.value;
        'content_id' = $oOCD.CONTENTID.value
        'grading_period_id' = $oOCD.GRADING_PERIODID.value;
        'date_created' = $oOCD.DATES.CREATED.value;
        'date_updated' = $oOCD.DATES.CREATED.value;
        'date_due' = $oOCD.DATES.DUE.value;
        'title' = $oOCD.TITLE.value;
        'position' = $oOCD.POSITION.value;
        'deleted' = $oOCD.DELETED.value;
        'weight' = $oOCD.WEIGHT.value;
        'points_possible' = $oOCD.POINTSPOSSIBLE.value;
        'is_visible' = $oOCD.ISVISIBLE.value;
        'visible_book' = $oOCD.VISIBLE_BOOK.value;
        'aggregation_model' = $oOCD.AGGREGATIONMODEL.value;
        'score_provider_handle' = $oOCD.SCORE_PROVIDER_HANDLE.value;
        'is_calculated' = $oOCD.ISCALCULATED.value;
        'calculation_type' = $oOCD.CALCULATIONTYPE.value;
        'is_scorable' = $oOCD.ISSCORABLE.value;
        'is_user_created' = $oOCD.ISUSERCREATED.value;
        'multiple_attempts' = $oOCD.MULTIPLEATTEMPTS.value;
        'is_delegated_grading' = $oOCD.IS_DELEGATED_GRADING.value;
        'is_anonymous_grading' = $oOCD.IS_ANONYMOUS_GRADING.value;
        'attempts' = @();
        'attempts_count' = 0;
      } 

      #endregion define outcome definition, aka ocd, aka column

      #region process content_id if it exists
      #at this level, it is the resource filename without extension, not the _id_ number
      if($ocd.content_id) {

        $cid = $ocd.content_id; 
        Write-Verbose ('Processing OCD Content: ' + $cid);
        
        [xml]$contentXML = Get-ZipEntryString -zipPath $bbaFile.Fullname -entryPath "$cid.dat";
        
        if($contentXML.CONTENT) {

          $content = $contentXML.CONTENT;
          if($content.Count -gt 1) { Write-Host ('WARNING: multiple content items in content file: ' + $cid) -ForegroundColor Yellow; }

          $co = [pscustomobject][ordered]@{
            'resource' = $cid;
            'id' = $content.id;
            'body' = $content.BODY.InnerText -replace $fix_comments;
            'content_handler' = $content.CONTENTHANDLER.value;
            'date_created' = $content.DATES.CREATED.value;
            'date_updated' = $content.DATES.UPDATED.value;
            'date_start' = $content.DATES.START.value;
            'date_end' = $content.DATES.end.value;
            'files' = '';
          }

          if($content.FILES) { $co.files = $content.FILES.OuterXml; } # may have to be dealt with later

          $content_items[$cid] = $co;

        } else {

          Write-Host ('Empty OCD Content file skipped: ' + "$cid.dat") -ForegroundColor Yellow;

        } #if CONTENT

      } #if content_id

      #endregion process content_id if it exists

      #region process outcomes
      $oOCD.OUTCOMES.OUTCOME |
        ForEach-Object {

        if(-not $_) { return; }

          $oOC = $_;

          $oc_id = $oOC.id;
          $cm_id = $oOC.COURSEMEMBERSHIPID.value;
          $sid = if($memberships.ContainsKey($cm_id)) { $memberships[$cm_id].user_id; }
          $oc_exempt = $oOC.EXEMPT.value;

          #region process attempts aka responses
          $oOC.ATTEMPTS.ATTEMPT |
            ForEach-Object {

              if(-not $_) { return; }

              $oAT = $_;
              $at_id = $oAT.id;

              if($memberships.ContainsKey($cm_id)) {
                $t_uid = $memberships[$cm_id].user_id; 
                $t_name = $roster[$t_uid].name; 
              } else { 
                $t_uid = $dummyUserId;
                $t_name = 'unknown';
              }

              #build an attempt object
              $ato = [pscustomobject][ordered]@{
                'id' = $oAT.id;
                'ocd_id' = $ocd_id;
                'oc_id' = $oc_id;
                'oc_exempt' = $oc_exempt;
                'cm_id' = $cm_id; 
                'user_id' = $t_uid;
                'user_name' = $t_name;
                'result_object_id' = $oAT.RESULTOBJECTID.value;
                'score' = $oAT.SCORE.value;
                'grade' = $oAT.GRADE.value;
                'status' = $oAT.STATUS.value;
                'date_added' = $oAT.DATEADDED.value;
                'date_attempted' = $oAT.DATEATTEMPTED.value;
                'date_first_gradeded' = $oAT.DATEFIRSTGRADEDED.value;
                'date_last_graded' = $oAT.DATELASTGRADED.value;
                'external_ref' = $oAT.EXTERNALREF.value;
                'student_comments' = $oAT.STUDENTCOMMENTS.InnerText -replace $fix_comments;
                'instructor_comments' = $oAT.INSTRUCTORCOMMENTS.InnerText -replace $fix_comments;
                'instructor_notes' = $oAT.INSTRUCTORNOTES.InnerText -replace $fix_comments;
                'comment_is_public' = $oAT.COMMENTISPUBLIC.InnerText -replace $fix_comments;
                'latest' = $oAT.LATEST.value;
                'exempt' = $oAT.EXEMPT.value;
                'group_attempt_id' = $oAT.GROUP_ATTEMPT_ID.value;
                'student_submission' = $oAT.STUDENTSUBMISSION.Value;
                'activity_counts' = $oAT.ACTIVITY_COUNTS.COUNT;
                'show_staged_feedback' = $oAT.SHOW_STAGED_FEEDBACK;
                'attempt_file_count' = ($attemptFiles[$at_id]).Count;
              }

              $attempt_items[$at_id] = $ato;

              #if latest, add to roster
              if($ato.latest -and $sid -and $roster.ContainsKey($sid)) {
                $roster[$sid].grades[$ocd_id] = $at_id;
              }
              
              $ocd.attempts += $at_id;
              $ocd.attempts_count++;

            } #ATTEMPTS.ATTEMPT

            #endregion process attempts aka responses

        } #OUTCOMES.OUTCOME;

        #end region process outcomes

        $outcome_items[$ocd_id] = $ocd;

    } #GRADEBOOK.OUTCOMEDEFINITIONS.OUTCOMEDEFINITION

Write-Progress $p3Name -Id 3 -Completed;

#region categories

$categories = @{};

$gbXml.GRADEBOOK.CATEGORIES.CATEGORY |
  ForEach-Object {

    if(-not $_) { return; }

    $categories[$_.id] = @{
      'id' = $_.id;
      'title' = $_.TITLE.value;
      'description' = $_.DESCRIPTION;
      'is_userdefined' = $_.ISUSERDEFINED.value;
      'is_calculated' =  $_.ISCALCULATED.value;
      'is_scorable' = $_.ISSCORABLE.value;
    }

  }

'Categories: ' + $categories.Count

#endregion categories

#region scales
$scales = @{};

$gbXml.GRADEBOOK.SCALES.SCALE |
  ForEach-Object {

    if(-not $_) { return; }

    $sid = $_.id;

    $scales[$sid] = [pscustomobject][ordered]@{
      'id' = $_.id;
      'description' = $_.DESCRIPTION.TEXT;
      'is_user_defined' = $_.ISUSERDEFINED.value;
      'is_tabular_scale' = $_.ISTABULARSCALE.value;
      'is_percentage' = $_.ISPERCENTAGE.value;
      'is_numeric' = $_.ISNUMERIC.value;
      'type' = $_.TYPE.value;
      'version' = $_.VERSION.value;
      'symbols' = $_.SYMBOLS.SYMBOL |
        ForEach-Object { 
          if(-not $_) { return; }
          @{ 
            $_.id = [pscustomobject][ordered]@{
              'id' = $_.id;
              'title' = $_.TITLE.value;
              'lowerbound' = $_.LOWERBOUND.value;
              'upperbound' = $_.UPPERBOUND.value;
              'abs_translation' = $_.ABSOLUTETRANSLATION.value;
            }
          }

        }

    }
  }
#endregion scales

#region grading_periods

  $gradingPeriods = $gbXml.GRADEBOOK.GRADING_PERIODS;
  if($gradingPeriods) {
    Write-Host 'This course has GRADING_PERIODS defined!' -ForegroundColor Yellow;
  }

#endregion grading_periods

#region formula

$formulae = @{};

$gbXml.GRADEBOOK.FORMULAE.FORMULA |
  ForEach-Object {

    if(-not $_) { return; }

    $gid = $_.GRADABLE_ITEM_ID.value;

    $formulae[$gid] = [pscustomobject][ordered]@{
      'id' = $_.id;
      'gradable_item_id' = $gid
      'jason' = ConvertFrom-Json $_.JSON_TEXT;
      'aliases' = $_.ALIASES.ALIAS |
        ForEach-Object { 
          if(-not $_) { return; }
          @{ 
            $_.name = [pscustomobject][ordered]@{
              'name' = $_.Name.value;
              'cid' = $_.CATEGORYID.value;
              'gid' = $_.GRADABLE_ITEM_ID.value;
            }
          }
        }

    } 
    
  } #GRADEBOOK.FORMULAE.FORMULA

#endregion formula

#endregion process_gradebook

Export-GBSummary; #write summary to title.csv

#region process_discussionboards
$fRes = $manXml.manifest.resources.resource | Where-Object { $_.type -eq $dbResType; };

##### Output table test of discusion board resources gathered
#$fRes | Format-Table -Property file,title,identifier,type
##### End Output test

$ePaths = $fRes.identifier;

'Discussion Boards: ' + $fRes.Count; # This assumes 1 discussion board per file

$discussionboards = @{}; #hash of discussion board objects by discussionboard.identifier

foreach ($file in $ePaths) {
  $fName = $file + '.dat';
  [xml]$dbXml = Get-ZipEntryString $bbaFile.Fullname -entryPath $fName; # Errors here because object returned isn't always filename (due to nested xml)

  if(-not $dbXml.FORUM) { throw 'Failed to get discussion board info'; }

  $dbXml.FORUM | ForEach-Object {

    $discussionboards[$_.id] = [pscustomobject][ordered]@{
      'id' = $_.id;
      'conferenceid' = $_.conferenceid.value;
      'title' = $_.title.value;
      'description' = $_.description.text;
    }
  }
}

$manXml.manifest.resources.resource | 
  Where-Object { $_.type -eq $dbResType; } |
    ForEach-Object {

      $base = $_.base;
      $fFile = $_.file | Select-Object -First 1;

      [xml]$fSubXml = Get-ZipEntryString $bbaFile.FullName -entryPath $fFile;

      #$fCount = ($fSubXml.FORUM).Count;
      #$fIndex = 0;
      #$usedFilenames = @();
      #$discussionboards = @{}; #hash of discussion board objects by discussionboard.identifier


      $fSubXml.FORUM | ForEach-Object {
        
        $fid = $_.id;
        $msgs = @{};

        #$mThreadMsgs = $fSubXml.FORUM.MESSAGETHREADS.MSG | ForEach-Object {
        $fSubXml.FORUM.MESSAGETHREADS.MSG | ForEach-Object {
          $mid = $_.id;

          $msg = [pscustomobject][ordered]@{
            'msg_id' = $_.id;
            'forum_id' = $_.FORUMID.value;
            'message_author' = $_.POSTEDNAME.value;
            'message_user_id' = $_.USERID.value;
            'message_title' = $_.TITLE.value;
            'message_text' = $_.MESSAGETEXT.TEXT;
            'date_created' = $_.DATES.CREATED.value;
            'date_updated' = $_.DATES.UPDATED.value;
            'date_last_edit' = $_.LASTEDITDATE.value;
            'date_last_post' = $_.LASTPOSTDATE.value;
          };

          if($mid -ne $null) {
            if($msgs.ContainsKey($mid)) {
              $msgs[$mid] += $msg;
            } else {
              $msgs[$mid] = $msg;
            }
          }
        };


        $fo = [pscustomobject][ordered]@{
          'base' = $base;
          'forum_id' = $_.id;
          'conference_id' = $_.CONFERENCEID.value;
          'title' = $_.TITLE.value;
          'uuid' = $_.UUID.value;
          'description' = $_.DESCRIPTION.TEXT;
          'messages' = $msgs;
        }

        #if($discussionboards.ContainsKey($fid)) {
        #  $discussionboards[$fid] += $fo;
        #} else {
        $discussionboards[$fid] = $fo;
        #}
      }
    }

    #$discussionboards.Values | ConvertTo-Json;



#endregion process_discussionboards

#region Generate HTML

#region html parts

$oct_bp = @"
<div class="ocd_title" onclick="toggle('{0}')" title="Click to open/close." ><span id="{0}_toggle" >&#9658; </span>{1}</div>
"@;

$file_bp = '<div class="{0}"><b>File:</b> <a href="{1}">{2}</a> [{3}]</div>';

$attrib_bp = ' <span class="ocAttribute"> {0} </span>';

$htmlStart = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <title>$bbaName</title>
  <link rel="stylesheet" type="text/css" href="bbaExtractorEX.css" />
  <script type="text/javascript">

    var t_closed = '&#9658; ';
    var t_open = '&#9660; ';

    function toggle(eid) {
      var tid = eid + '_toggle';
      var oTitle = document.getElementById(eid);
      var oSpan = document.getElementById(tid);
      
      if(oTitle.style.display == 'block') {
        oTitle.style.display = 'none';
        oSpan.innerHTML = t_closed;
      } else {  
        oTitle.style.display = 'block';
        oSpan.innerHTML = t_open;
      }
    }


  </script>
</head>
<body>
<div class="page_title">Extracted BBArchive Course Info</div>
"@;

#endregion html parts

Push $htmlStart;

#region Course Info

Push-Comment ('CourseInfo: ' + $csRes.file) -tee;
Push-Comment ('Course: ' + $csXml.COURSE.TITLE.value) -tee;
Push '<div class="course_info">';
Push-DIV 'ci_title' '' $courseInfo.title;
Push '<div class="course_details">';
Push-DIV 'cd_archive' 'Archive:' $bbaFile.Name $bbaFile.FullName;
Push-DIV 'cd_exported' 'Exported:' ('{0}*' -f $manEntry.LastWriteTime) 'This is the LastWriteTime of the manifest file.';
Push-DIV 'cd_created' 'Course Created:' $courseInfo.date_created;
Push-DIV 'cd_updated' 'Course Updated:' $courseInfo.date_updated;
Push-DIV 'cd_start' 'Course Start:' $courseInfo.date_start;
Push-DIV 'cd_end' 'Course End:' $courseInfo.date_end;
Push '</div><!--course_details-->';
Push '</div><!--course_info-->';

#endregion Course Info


#region Roster HTML

Push '<div class="roster">';

Push @"
<div class="roster_title" onclick="toggle('_roster_')" title="Click to open/close." ><span id="roster_toggle">&#9658; </span>Course Roster</div>
"@;

Push '<div class="roster_body" id="_roster_">';

Push ($roster.Values | Where-Object { $_.username -and $_.role -ne 'STUDENT' } | Sort-Object role, username | ConvertTo-Html -Fragment);

Push ($roster.Values | Where-Object { $_.username -and $_.role -eq 'STUDENT' } | Sort-object username | ConvertTo-Html -Fragment);

Push '</div><!--roster_body-->';

Push '</div><!--roster-->';

#endregion Roster HTML


#region columns adk outcome definitions

Push '<div class="outcome_definitions">';

$progIndex = 0;
$progCount = $outcome_items.Count;

$p4Name = 'Rendering HTML';

$outcome_items.Values | 
  ForEach-Object {

    if(-not $_) { return; }

    $title = $_.title

    Write-Progress $p4Name -Status ('{0}/{1} - {2}' -f ++$progIndex, $progCount, $title) -PercentComplete ($progIndex/$progCount*100) -Id 4;

    Push-Comment ('OUTCOMEDEFINITION: ' + $_.id);
    Push '<div class="ocd">';

    $titleHTML = $_.title;

    if($_.is_visible -eq 'false') {
      $titleHTML += ($attrib_bp -f 'Not Visible');
    }

    if($_.is_calculated -eq 'true') {
      $titleHTML += ($attrib_bp -f 'Calculated');
    }

    Push ($oct_bp -f $_.id, $titleHTML);

    Push '<div class="ocd_info">';
    Push-DIV 'ocdi_due' 'Due:' $_.date_due;
    Push-DIV 'ocdi_possible' 'Possible:' $_.points_possible;
    Push-DIV 'ocdi_created' 'Created:' $_.date_created;
    Push-DIV 'ocdi_updated' 'Updated:' $_.date_updated;

    if($_.weight -gt 0) {
      Push-DIV 'ocdi_weight' 'Weight:' $_.weight;
    }

    if($_.score_provider_handle) {
      Push-DIV 'ocdi_provider' 'Score Provider:' $_.score_provider_handle;
    };

    Push '</div><!--ocd_info-->';

    if($_.content_id -and $content_items[$_.content_id]) {

      Push-DIV 'ocd_content' '' ($content_items[$_.content_id].body.trim() -replace $fix_content); 

    }

    Push ('<div class="outcomes" id="{0}">' -f $_.id);

    #below, passing an array of keys to a hash.. did not know I could even do that!
    $attempt_items[$_.attempts] | Sort-Object user_name | ForEach-Object {

      if(-not $_) { return; }

      $aid = $_.id;

      Push '<div class="oc">';

      Push-Comment ('USER: ' + $_.user_id);
      Push '<div class="user_info">';

      $uo = $roster[$_.user_id];

      Push-DIV 'ui_username' 'Username:' $uo.username;
      Push-DIV 'ui_student_id' 'StudentID:' $uo.student_id;
      Push-DIV 'ui_name' 'Name:' $uo.name;
      Push-DIV 'ui_role' 'Role:' $uo.role;

      Push '</div><!--user_info-->';

      Push-Comment ('ATTEMPT: ' + $aid);
      Push '<div class="at">'; 
      Push '<div class="at_details">';

      if(($_.score -replace '\.0$', '') -eq ($_.score -replace '\.0$', '')) {
        $score = $_.score;
      } else {
        $score = $_.score + '/' + $_.grade;
      }

      if($_.exempt -eq 'true') {
        $score = '<span class="atd_exempt">{0} [Exempt]</span>' -f $score;
      }

      Push-DIV 'atd_score' 'Score:' $score;
      Push-DIV 'atd_status' 'Status:' $_.status;
      Push-DIV 'atd_date' 'Date:' $_.date_attempted
      Push-DIV 'atd_latest' 'Latest:' $_.latest;

      Push '</div><!--at_details-->';

      #comments
      if($_.student_comments -or $_.instructor_comments -or $_.instructor_notes) {

      Push '<div class="at_comments">';
      
      if($_.student_comments) {
          Push-DIV 'atc_scom', 'S-Comment: ', $_.student_comments;
        }

        if($_.instructor_comments) {
          Push-DIV 'atc_icom' 'I-Comment: ' $_.instructor_comments;
        }

        if($_.instructor_notes) {
          Push-DIV 'atc_inote' 'I-Note: ' ($_.instructor_notes -replace $fix, ' ');
        }

        Push '</div><!--at_comments-->';

      } # comments

      #files
      if($attemptFiles.ContainsKey($aid)) {
  
        Push '<div class="at_files">';

        $attemptFiles[$aid] | ForEach-Object {

          $fdc = $_.date_created;
          if($_.date_updated) { $fdc += '/' + $_.date_updated; }
          Push ($file_bp -f 'atf', $_.extracted_path, $_.file_link_name, $fdc);

        }

        Push '</div><!--at_files-->';

      } #has files

      Push '</div><!--at-->'; 
      Push '</div><!--oc-->';

    } #$outcome_items.Values.attempts

    Push '</div><!--outcomes-->';
    Push '</div><!--ocd-->';

  } #$outcome_items.Values

Push '</div><!--outcome_definitions-->';

#region columns aka outcome definitions

#endregion

#region Discussion Boards
Push '<div class="discussion_boards">';

$discussionboards.Values | ForEach-Object {
  # Output Discussion Board Info
  #Push '<div class="discussion" id='
  #  Push $_.forum_id
  #Push '>'
  Push ('<div class="discussion">'-f $_.forum_id);
    Push '<div class="discussion_info">';
      #Push-DIV 'discussion_title' 'Discussion:' $_.title;
      $oct_bp = @"
<div class="ocd_title" onclick="toggle('{0}')" title="Click to open/close." ><span id="{0}_toggle" >&#9658; </span>{1}</div>
"@;
      Push ('<div class="discussion_title" onclick="toggle(''{0}'')" title="Click to open/close."><span id="{0}_toggle">&#9658; </span>Discussion: {1}</div>'-f $_.forum_id, $_.title);
      Push-DIV 'discussion_description' '' $_.description;
    Push '</div><!--discussion_info-->';

    # Output each message in board
    #Push '<div class="messages" id=f_'
    #  Push $_.forum_id
    #Push '>'
    Push ('<div class="messages" id="{0}">'-f $_.forum_id);
      $_.messages.Values | ForEach-Object {
        #Push '<div class="message" id=' 
        #  Push $_.msg_id 
        #Push '>' 
        Push ('<div class="message" id="{0}">'-f $_.msg_id);
        Push-DIV 'message_title' 'Message Title:' $_.message_title;
        Push '<div class="message_info">'
          Push-DIV 'message_author' 'Author:' $_.message_author;
          Push-DIV 'message_created' 'Created:' $_.date_created;
          Push-DIV 'message_updated' 'Updated:' $_.date_updated;
          Push-DIV 'date_last_edit' 'Last Edited:' $_.date_last_edit;
          Push-DIV 'date_last_post' 'Last Post:' $_.date_last_post;
        Push "</div>"
          Push-DIV 'message_text' '' $_.message_text;
        Push "</div>"
      }
    Push '</div><!--messages-->';
  Push '</div><!--discussion-->';
}

Push '</div><!--discussion_boards-->';
Push "</div>"

#endregion

#region Close HTML

Push @"
<div class=`"notes`">
  *Export time is based on the creation time of the manifest.<br/>
  *This HTML file was generated: $extracted<br/>
</div><!--notes-->
</body>
</html>
"@ -flush;

#endregion Close HTML

#endregion Generate HTML

Write-Progress $p4Name -Id 4 -Completed;

if($viewResult -and $PSVersionTable.PSEdition -eq 'Desktop' -and (Test-Path $htmlFile)) {
  Start-Process $htmlFile;
}

