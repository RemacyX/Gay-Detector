Dim gayPercentage, userName, birthYear, age, retry, sexualOrientation, location, gender, partners, relationshipStatus

Do
    retry = False
    userName = InputBox("Please enter your name:")
    If userName = "" Then
        msgbox "You didn't enter a name. Please try again.", vbExclamation, "Error"
        retry = True
    Else
        birthYear = InputBox("Please enter your birth year:")
        age = Year(Now) - birthYear
        gayPercentage = Int((100 - 1 + 1) * Rnd + 1)
        If gayPercentage < 50 Then
            If gayPercentage <= 25 Then
                msg = "You are not gay. You are straight."
            Else
                msg = "You are not gay but you have slight gay tendencies."
            End If
        Else
            If gayPercentage <= 75 Then
                msg = "You are gay but you are in the closet."
            Else
                msg = "You are openly gay."
            End If
        End If
        sexualOrientation = InputBox("Please select your sexual orientation (straight/gay/bisexual/other)")
        If sexualOrientation = "other" Then
            msgbox "Please select a valid option (straight/gay/bisexual)"
            retry = True
        Else
            location = InputBox("Please enter your location (City, Country)")
            gender = InputBox("Please select your gender (male/female/other)")
            partners = InputBox("Please enter the number of sexual partners you have had")
            relationshipStatus = InputBox("Please enter your relationship status (single/married/in a relationship/other)")
            msgbox "Welcome " & userName & " !" & vbNewLine & "You are " & age & " years old." & vbNewLine & "Your sexual orientation is " & sexualOrientation & vbNewLine & "Your location is " & location & vbNewLine & "Your gender is " & gender & vbNewLine & "You have had " & partners & " sexual partners" & vbNewLine & "Your relationship status is " & relationshipStatus & vbNewLine & vbNewLine & "According to our Gay Detector, your gay percentage is: " & gayPercentage & "%" & vbNewLine & vbNewLine & msg, vbInformation, "Gay Detector"
        End If
    End If
Loop Until retry = False

msgbox "Script by Remacy" & vbNewLine & "This script is open-source and does not log any information", vbInformation, "Credits"
