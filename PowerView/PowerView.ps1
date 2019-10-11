<#
PowerView - An OSRS player stat viewer
February 12, 2018
By Daniel Shaw

This script can either be run as a standalone to view stats, or in an interactive environment:

#Export selected rows in Out-GridView to a CSV file
PS> .\PowerView.ps1 | Export-Csv -Path .\stats.csv

#Output selected rows to JSON
PS> .\PowerView.ps1 | ConvertTo-Json
#>
param([string]$Name)

Function To-Name($Index) {
    switch($Index) {
        0  { "Total"        }
        1  { "Attack"       }
        2  { "Defence"      }
        3  { "Strength"     }
        4  { "Hitpoints"    }
        5  { "Ranged"       }
        6  { "Prayer"       }
        7  { "Magic"        }
        8  { "Cooking"      }
        9  { "Woodcutting"  }
        10 { "Fletching"    }
        11 { "Fishing"      }
        12 { "Firemaking"   }
        13 { "Crafting"     }
        14 { "Smithing"     }
        15 { "Mining"       }
        16 { "Herblore"     }
        17 { "Agility"      }
        18 { "Thieving"     }
        19 { "Slayer"       }
        20 { "Farming"      }
        21 { "Runecraft"    }
        22 { "Hunter"       }
        23 { "Construction" }
    }
}

Function To-Stat($Line, $Index) {
    if($Line.Split(",").Length -ige 3) {
        $content = $Line.Split(",");
        return [PSCustomObject]@{
            Skill = To-Name($Index)
            Rank = [int]("{0:N0}" -f [int]$content[0])
            Level = [int]$content[1]
            Experience = [int]("{0:N0}" -f [int]$content[2])
        }
    }
}

$url = "services.runescape.com/m=hiscore_oldschool/index_lite.ws?player="
try {
    $request = Invoke-WebRequest -Uri ("{0}{1}" -f $url, $Name)
} catch {
    Write-Error "`"$Name`" is not a valid OSRS display name." -Category InvalidArgument -ErrorId "1"
    return
} 

$raw_data = $request.Content.Split("`n")
#Will contain the completely parsed stats.
$data = @()
$index = 0
foreach($i in $raw_data) {
    #This excludes all non-skill based lines because honestly, who cares.
    if(-Not [string]::IsNullOrEmpty($i.Trim())) {
        $data += (To-Stat -Line $i -Index $index);
        $index++
    }
}
$data