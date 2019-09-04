Function Script-Module-SetPassword {
    param (
        [int]$PasswordLength
    )
    # ============
    # Declarations
    # ============
    $global:Password = $null
    $global:VerificationCode = $null
    $Vowels = "aeiouy"
    $Consonants = "bcdfghjklmnpqrstvwxz"
    $Numbers = "0123456789" 
    $Specials = "!"
    $PreviousCharIsVowel = $true
    $DoubleConsonant = $false
    If (!$PasswordLength) {
        $PasswordLength = 8
    }
    # =========
    # Execution
    # =========
    # -------------------------------------------------------------
    # Defining 2 positions for a random digit and special character
    # -------------------------------------------------------------
    $Position1 = ([int]($PasswordLength / 2)) + (Get-Random ([int]($PasswordLength / 2)))
    $Position2 = ($Position1 - 1) - (Get-Random ([int]($Position1 / 2)))
    If (Get-Random 2) {
        $PositionSpecial = $Position1
        $PositionNumber = $Position2
    }
    Else {
        $PositionNumber = $Position1
        $PositionSpecial = $Position2
    }
    # ------------------------------------------------------------------------------------------
    # Generating a random password by iterating through the total length to generate each letter
    # the reason for this is to enforce a maximum of 2 consonants after each other
    # ------------------------------------------------------------------------------------------
    ForEach ($Counter in 0..($PasswordLength - 1)) {
        # ----------------------------------------------------------------
        # Generating a number when the position defined earlier is reached
        # ----------------------------------------------------------------
        If ($Counter -eq $PositionNumber) {
            $GeneratedChar = $Numbers[(Get-Random $Numbers.Length)]
        }
        # ---------------------------------------------------------------------------
        # Generating a special character when the position defined earlier is reached
        # ---------------------------------------------------------------------------
        ElseIf ($Counter -eq $PositionSpecial) {
            $GeneratedChar = $Specials[(Get-Random $Specials.Length)]
        }
        Else {
            # ---------------------------------------------------------
            # Generate a consonant if the previous character is a vowel
            # ---------------------------------------------------------
            If ($PreviousCharIsVowel -eq $true) {
                $GeneratedChar = ($Consonants[(Get-Random $Consonants.Length)])
            }
            Else {
                # ----------------
                # Generate a vowel
                # ----------------
                $GeneratedChar = ($Vowels[(Get-Random $Vowels.Length)])
            }
            # ----------------------------------
            # Randomly set capital on the letter
            # ----------------------------------
            If ((Get-Random 2) -eq 1) {
                $GeneratedChar = [char]([int]$GeneratedChar - 32)
            }
            # ----------------------------------------------------------------------------------------------
            # If the previous character is a vowel or not a double consonant, then set the state accordingly
            # ----------------------------------------------------------------------------------------------
            If (!$DoubleConsonant -AND $PreviousCharIsVowel -AND $Counter -gt 0) {
                If (Get-Random 2) {
                    $PreviousCharIsVowel = !$PreviousCharIsVowel
                }
                Else {
                    $DoubleConsonant = $true
                }
            }
            Else {
                # --------------------------------------------------------------
                # If a second consonant has been generated, then we need a vowel
                # --------------------------------------------------------------
                $PreviousCharIsVowel = !$PreviousCharIsVowel
                $DoubleConsonant = $false
            }
        } 
        $global:Password += $GeneratedChar
    }
    # ---------------------------------------------------------------------------------------------------------------
    # Generate a verification code to transfer the password to the user (this is required due to false impersonation)
    # ---------------------------------------------------------------------------------------------------------------
    ForEach ($Counter in 0..2) {
        $global:VerificationCode += $Numbers[(Get-Random 10)]
    }
}
