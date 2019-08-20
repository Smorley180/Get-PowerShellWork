function Is-FizzBuzz {
    param (
        # Numbers to iterate over
        [Parameter(Mandatory = $false,
                    ValueFromPipeline = $true)]
        $Numbers = 1..100,
        # First value to check, default = 3
        [Parameter(Mandatory = $false)]
        [int]
        $FirstNumber = 3,
        # Second value to check, default = 5
        [Parameter(Mandatory = $false)]
        [int]
        $SecondNumber = 5
    )
    $Numbers | ForEach-Object {
        if (($_ % $FirstNumber -eq 0) -and ($_ % $SecondNumber -eq 0)) {
            Write-Output "FizzBuzz"
        }
        elseif ($_ % $FirstNumber -eq 0) {
            Write-Output "Fizz"
        }
        elseif ($_ % $SecondNumber -eq 0) {
            Write-Output "Buzz"
        }
        else {
            Write-Output $_
        }
    }
}