#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Persistent
SetBatchLines, -1  ; Run script at full speed
DetectHiddenWindows, On

; 🔹 Keyboard Shortcut to Stop Script
^Esc::ExitApp  ; Press Ctrl + Esc to stop the script

; 🔹 Set Variables
LinkedInJobsURL := "https://www.linkedin.com/jobs"
ResumePath := "C:\Users\Jayme\Documents\Project_Manager_IT_Resume_2025.pdf"
GoogleSheetsForm := "https://docs.google.com/forms/d/e/1FAIpQLSfzXyW4XzT9uY5jDVm1_A5s4x9YqN0c1VpoQWv7nnJdMcmKhg/viewform"

GmailUser := "jaymem89.jm@gmail.com"
GmailPass := "sgpd qnql ahzi ceje"

MsgBox, Script Started! 🚀

Loop
{
    MsgBox, Fetching Jobs...
    jobList := FetchJobsFromLinkedIn()
    
    If (jobList = "Error fetching jobs")
    {
        MsgBox, ❌ Job fetching failed! Check network connection.
        ExitApp
    }

    ; 🔹 Process Each Job Application
    Loop, Parse, jobList, "`n"
    {
        jobDetails := StrSplit(A_LoopField, "|")
        JobTitle := jobDetails[1]
        CompanyName := jobDetails[2]
        JobURL := jobDetails[3]

        MsgBox, Found Job: %JobTitle% at %CompanyName%`nURL: %JobURL%

        ; Extract Recruiter Email (if available)
        RecruiterEmail := ExtractRecruiterEmail(JobURL)

        MsgBox, Recruiter Email: %RecruiterEmail%

        ; Apply for the job
        ApplyToJob(JobTitle, CompanyName, JobURL)

        ; Log Job in Google Sheets
        LogJobInGoogleSheets(JobTitle, CompanyName, JobURL, RecruiterEmail)

        ; Schedule Follow-Up Email in 4 Days
        SetTimer, SendFollowUpEmail, -345600000  ; 4 days delay
    }

    MsgBox, Waiting 1 minute before fetching jobs again...
    Sleep, 60000  ; Wait 1 minute before fetching new jobs again
}

return

; 🔹 Function: Fetch Jobs from LinkedIn
FetchJobsFromLinkedIn()
{
    try {
        Url := "https://www.linkedin.com/jobs/search/?keywords=Cybersecurity&location=Remote"
        HttpObj := ComObjCreate("MSXML2.XMLHTTP")
        HttpObj.Open("GET", Url, False)
        HttpObj.Send()

        JobData := HttpObj.ResponseText
        If JobData =
            return "Error fetching jobs"
        
        return ParseJobListings(JobData)
    }
    catch
    {
        return "Error fetching jobs"
    }
}

; 🔹 Function: Extract Recruiter Email
ExtractRecruiterEmail(JobURL)
{
    try {
        HttpObj := ComObjCreate("MSXML2.XMLHTTP")
        HttpObj.Open("GET", JobURL, False)
        HttpObj.Send()

        PageContent := HttpObj.ResponseText
        If RegExMatch(PageContent, "([\w._%+-]+@[\w.-]+\.[a-zA-Z]{2,4})", Match)
            return Match1
        else
            return "No email found"
    }
    catch
    {
        return "Error extracting email"
    }
}

; 🔹 Function: Apply to Job (Simulated API Call)
ApplyToJob(JobTitle, CompanyName, JobURL)
{
    try {
        MsgBox, Applying for job: %JobTitle% at %CompanyName%
        return "Applied Successfully"
    }
    catch
    {
        return "Application Failed"
    }
}

; 🔹 Function: Log Job in Google Sheets
LogJobInGoogleSheets(JobTitle, CompanyName, JobURL, RecruiterEmail)
{
    try {
        FormURL := GoogleSheetsForm "?entry.1234567890=" . JobTitle . "&entry.0987654321=" . CompanyName . "&entry.1122334455=" . JobURL . "&entry.2233445566=" . RecruiterEmail
        MsgBox, Logging Job to Google Sheets...`n%FormURL%
        Run, %FormURL%
        Sleep, 5000
        return "Logged Successfully"
    }
    catch
    {
        return "Logging Failed"
    }
}

; 🔹 Function: Send Follow-Up Email via Gmail SMTP
SendFollowUpEmail()
{
    global JobTitle, CompanyName, JobURL, RecruiterEmail, GmailUser, GmailPass, ResumePath

    if (RecruiterEmail != "No email found")
    {
        Msg := "Subject: Follow-Up on " . JobTitle . "`r`n"
        Msg .= "To: " . RecruiterEmail . "`r`n"
        Msg .= "From: Jayme Montgomery <" . GmailUser . ">`r`n"
        Msg .= "`r`n"
        Msg .= "Hello," . "`r`n`r`n"
        Msg .= "I hope you’re doing well. I wanted to follow up on my application for the " . JobTitle . " position at " . CompanyName . ".`r`n`r`n"
        Msg .= "I’m very excited about the opportunity and wanted to see if there are any updates regarding my application.`r`n`r`n"
        Msg .= "Please let me know if there’s any additional information I can provide.`r`n`r`n"
        Msg .= "Best regards," . "`r`n"
        Msg .= "Jayme Montgomery" . "`r`n"
        Msg .= "469-604-8709" . "`r`n"
        Msg .= "jaymem89.jm@gmail.com"

        MsgBox, Sending Follow-Up Email to %RecruiterEmail%

        ; Send email using Gmail SMTP
        Run, PowerShell -Command "& {Send-MailMessage -SmtpServer smtp.gmail.com -Port 587 -UseSsl -Credential (New-Object PSCredential ('" . GmailUser . "', (ConvertTo-SecureString '" . GmailPass . "' -AsPlainText -Force))) -From '" . GmailUser . "' -To '" . RecruiterEmail . "' -Subject 'Follow-Up on " . JobTitle . "' -Body '" . Msg . "' -Attachments '" . ResumePath . "' }",, Hide
        return "Follow-up Email Sent"
    }
    else
    {
        MsgBox, No recruiter email found, skipping follow-up.
        return "No recruiter email found, skipping follow-up"
    }
}

; 🔹 Function: Parse Job Listings from LinkedIn's HTML
ParseJobListings(JobData)
{
    jobList := ""
    Loop, Parse, JobData, "`n"
    {
        If RegExMatch(A_LoopField, "<a.*?href=""(https://www.linkedin.com/jobs/view/\d+)"".*?>(.*?)</a>", Match)
        {
            JobURL := Match1
            JobTitle := Match2
            jobList .= JobTitle . "|Unknown Company|" . JobURL . "`n"
        }
    }
    Return jobList
}