#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testFixedAddressWorksheet = @(
        [PSCustomObject]@{
            Name          = 'Test'
            ddns_hostname = 'Test'
            disable       = $false
            enable_ddns   = $true
            ipv4addr      = '0.0.0.0'
            mac           = '00:00:00:00:00:00'
            match_client  = 'RESERVED'
            overwrite     = $null
            comment       = 'Pester test 1'
        }
        [PSCustomObject]@{
            Name          = 'Test2'
            ddns_hostname = 'Test2'
            disable       = $false
            enable_ddns   = $true
            ipv4addr      = '10.10.10.1'
            mac           = '00:00:00:00:00:AA'
            match_client  = 'MAC_ADDRESS'
            overwrite     = $null
            comment       = 'Pester test 2'
        }
    )
    $testRemoveWorksheet = @(
        [PSCustomObject]@{
            Name          = 'Test'
            ddns_hostname = 'Test'
            ipv4addr      = $null
            mac           = $null
        }
        [PSCustomObject]@{
            Name          = 'Test2'
            ddns_hostname = 'Test2'
            ipv4addr      = '10.10.10.1'
            mac           = '00:00:00:00:00:AA'
        }
    )
    
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (GBB)'
        Environment = 'Test'
        MailTo      = 'BobLeeSwagger@shooter.net'
        ImportFile  = New-Item 'TestDrive:/IPAM.xlsx' -ItemType File
        LogFolder   = New-Item 'TestDrive:/Log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    $MailAdminParams = {
        ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }
    $MailUsersParams = {
        ($To -eq $testParams.MailTo) -and ($Priority -eq 'High') -and ($Subject -like 'FAILURE - Incorrect input')
    }

    Mock Get-IpamFixedAddressHC
    Mock New-IpamFixedAddressHC
    Mock Remove-IpamObjectHC
    Mock Restart-IpamServiceHC
    Mock Resolve-DnsName
    Mock Import-Excel
    Mock Export-Excel
    Mock Invoke-WebRequest
    Mock Send-MailHC -RemoveParameterValidation Attachments
    Mock Write-EventLog
    Mock Test-Connection
}
Describe 'Import file' {
    Context 'send an error mail to the admin when' {
        It 'the file is not found' {
            $testNewParams = Copy-ObjectHC -Name $testParams
            $testNewParams.ImportFile = 'NonExisting.xlsx'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Import file '$ImportFile' not found*")
            }
        } 
        It "the worksheet 'FixedAddress' is missing" {
            Mock Import-Excel
            Mock Import-Excel {
                throw 'Worksheet not found'
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like '*Worksheet not found*')
            }
        } 
        It "the worksheet 'Remove' is missing" {
            Mock Import-Excel
            Mock Import-Excel {
                throw 'Worksheet not found'
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like '*Worksheet not found*')
            }
        } 
    }
}
Describe 'FixedAddress worksheet' {
    BeforeAll {
        Mock Import-Excel
    }
    Context 'send an error mail to the users when' {
        It "an incorrect 'ipv4addr' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.ipv4addr = '10.10.10.300'
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*ipv4addr*not a valid IP*')
            }
        } 
        It "an incorrect 'mac' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.ipv4addr = '10.10.10.2'
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1.mac = '00:00:00:00:00'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*mac*not*valid*')
            }
        } 
        It "'enable_ddns' is TRUE but 'ddns_hostname' is missing" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.enable_ddns = $true
                $testF1.ddns_hostname = $null
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*ddns_hostname*')
            }
        } 
        Context 'a duplicate value is found for' {
            It 'mac' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.ipv4addr = '10.10.10.1'
                    $testF1.match_client = 'MAC_ADDRESS'
                    $testF1.mac = '00:00:00:00:AA'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF2.ipv4addr = '10.10.10.2'
                    $testF2.match_client = 'MAC_ADDRESS'
                    $testF2.mac = '00:00:00:00:AA'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*mac*used multiple times*')
                }
            } 
            It 'ipv4addr' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.ipv4addr = '10.10.10.1'
                    $testF1.match_client = 'MAC_ADDRESS'
                    $testF1.mac = '00:00:00:00:AA'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF2.ipv4addr = '10.10.10.1'
                    $testF2.match_client = 'MAC_ADDRESS'
                    $testF2.mac = '00:00:00:00:BB'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*ipv4addr*used multiple times*')
                }
            } 
            It 'name' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.name = 'kiwi'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testFixedAddressWorksheet[1]
                    $testF2.name = 'kiwi'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*name*used multiple times*')
                }
            } 
            It 'ddns_hostname' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.ddns_hostname = 'kiwi'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testFixedAddressWorksheet[1]
                    $testF2.ddns_hostname = 'kiwi'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*ddns_hostname*used multiple times*')
                }
            } 
        }
        Context 'a mandatory property is missing' {
            BeforeAll {
                $TestCases = @(
                    'name',
                    'ipv4addr',
                    'match_client'
                ).ForEach( { @{Name = $_ } })
            }

            It '<Name>' -TestCases $TestCases {
                Param (
                    $Name
                )
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.PSObject.Properties.Remove($Name)
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.$Name = $null
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 2 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*mandatory property '$Name' is missing*")
                }
                Should -Invoke Write-EventLog -Exactly 2 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            } 
        }
        Context 'a boolean property is not TRUE, FALSE or NULL' {
            BeforeAll {
                $TestCases = @(
                    'enable_ddns',
                    'overwrite',
                    'disable'
                ).ForEach( { @{Name = $_ } })
            }
            It '<Name>' -TestCases $TestCases {
                Param (
                    $Name
                )

                foreach ($testVal in @($true, $false, $null)) {
                    Mock Import-Excel
                    Mock Import-Excel {
                        $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                        $testF1.$Name = $testVal
                        $testF1
                    } -ParameterFilter {
                        $WorksheetName -eq 'FixedAddress'
                    }

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 0 -ParameterFilter {

                        (&$MailUsersParams) -and ($Message -like "*boolean property*")
                    }
                }

                Mock Import-Excel
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.$Name = 'Wrong'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*boolean property*")
                }
            } 
        }
    }
    Context 'string properties' {
        It 'leading and trailing spaces are removed' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.Comment = ' TEST  '
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            $FixedAddressWorksheet.Comment | Should -Be  'TEST'
        } 
    }
    Context 'mac' {
        It "is converted from '00-00-00-00-00-00' to '00:00:00:00:00:00'" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1.mac = '00-00-00-00-00-00'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            $FixedAddressWorksheet.mac | Should -Be  '00:00:00:00:00:00'
        } 
        It "mac is converted to lower case" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1.mac = 'AA-BB-00-00-00-00'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            $FixedAddressWorksheet.mac | Should -Be  'aa:bb:00:00:00:00'
        } 
    }
    Context 'match_client' {
        Context 'send an error mail to the user when' {
            It 'the value for match_client is unknown' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'wrong'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*The field 'match_client' contains the unsupported value 'wrong'*")
                }
            }
            It 'MAC_ADDRESS is set but the mac is missing' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'MAC_ADDRESS'
                    $testF1.mac = $null
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*The field 'mac' is missing*When 'match_client' is set to 'MAC_ADDRESS' the field 'mac' is mandatory*")
                }
            } 
            It "RESERVED is set but a mac that is not '00:00:00:00:00:00' is given" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'RESERVED'
                    $testF1.mac = 'AA:AA:AA:AA:AA:AA'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*The field 'mac' cannot be set to 'AA:AA:AA:AA:AA:AA' when 'match_client' is set to 'RESERVED'*")
                }
            } 
        }
        Context 'send no error mail when' {
            It "RESERVED is set with mac '00:00:00:00:00:00'" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'RESERVED'
                    $testF1.mac = '00:00:00:00:00:00'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 0 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*match_client*")
                }
            } 
            It 'RESERVED has no mac' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'RESERVED'
                    $testF1.mac = $null
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 0 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*Tmatch_client*")
                }
            } 
            It "MAC_ADDRESS is set with mac '00:00:00:00:00:00'" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'MAC_ADDRESS'
                    $testF1.mac = '00:00:00:00:00:00'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 0 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*Tmatch_client*")
                }
            } 
        }
        Context "set match_client to 'RESERVED' and mac to '00:00:00:00:00:00' when" {
            It 'RESERVED is set and the mac is blank' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'RESERVED'
                    $testF1.mac = $null
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                $IpamBody.match_client | Should -Be 'RESERVED'
                $IpamBody.mac | Should -Be '00:00:00:00:00:00'
            } 
            It "MAC_ADDRESS is set and the mac is '00:00:00:00:00:00'" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.match_client = 'MAC_ADDRESS'
                    $testF1.mac = '00:00:00:00:00:00'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                $IpamBody.match_client | Should -Be 'RESERVED'
                $IpamBody.mac | Should -Be '00:00:00:00:00:00'
            } 
        }
    }
    Context 'ipv4addr' {
        Context 'send an error mail to the user when' {
            Context "ipv4addr starts with 'func:'" {
                It "and match_client is not 'MAC_ADDRESS'" {
                    Mock Import-Excel {
                        $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                        $testF1.match_client = 'RESERVED'
                        $testF1.ipv4addr = 'func:nextavailableip:10.20.32.0/24'
                        $testF1
                    } -ParameterFilter {
                        $WorksheetName -eq 'FixedAddress'
                    }

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailUsersParams) -and ($Message -like "*ipv4addr*func*match_client*RESERVED*")
                    }
                } 
                It "and mac is '00:00:00:00:00:00'" {
                    Mock Import-Excel {
                        $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                        $testF1.match_client = 'MAC_ADDRESS'
                        $testF1.mac = '00:00:00:00:00:00'
                        $testF1.ipv4addr = 'func:nextavailableip:10.20.32.0/24'
                        $testF1
                    } -ParameterFilter {
                        $WorksheetName -eq 'FixedAddress'
                    }

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailUsersParams) -and ($Message -like "*ipv4addr*func*match_client*")
                    }
                } 
                It "and mac is blank" {
                    Mock Import-Excel {
                        $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                        $testF1.match_client = 'MAC_ADDRESS'
                        $testF1.mac = $null
                        $testF1.ipv4addr = 'func:nextavailableip:10.20.32.0/24'
                        $testF1
                    } -ParameterFilter {
                        $WorksheetName -eq 'FixedAddress'
                    }

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailUsersParams) -and ($Message -like "*ipv4addr*func*match_client*")
                    }
                } 
            }
        }
    }
    Context 'ddns_hostname' {
        Context 'send an error mail to the user when' {
            It "ddns_hostname contains a space or a dot to avoid FQDN" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.ddns_hostname = 'test name'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.ddns_hostname = 'test.domain.net'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 2 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*ddns_hostname*")
                }
            } 
        }
    }
    Context 'name' {
        Context 'send an error mail to the user when' {
            It "name contains a space or a dot" {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.name = 'test name'
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.name = 'test.domain.net'
                    $testF1
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 2 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like "*name*")
                }
            } 
        }
    }
}
Describe 'Remove worksheet' {
    BeforeAll {
        Mock Import-Excel
    }
    Context 'send an error mail to the users when' {
        It "an incorrect 'ipv4addr' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.ipv4addr = '10.10.10.300'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*ipv4addr*not a valid IP*')
            }
        } 
        It "an incorrect 'mac' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.mac = '00:00:ddd:zzz:00'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*mac*not*valid*')
            }
        } 
        It "an incorrect 'ddns_hostname' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.ddns_hostname = 'host.contoso.com'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*ddns_hostname*')
            }
        } 
        It "a blank 'mac' with '00:00:00:00:00:00' is given" {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.mac = '00:00:00:00:00:00'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*mac*not*allowed*')
            }
        } 
        Context 'a duplicate value is found for' {
            It 'mac' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF1.mac = '00:00:00:00:AA'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF2.mac = '00:00:00:00:AA'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*mac*used multiple times*')
                }
            } 
            It 'ipv4addr' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF1.ipv4addr = '10.10.10.1'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF2.ipv4addr = '10.10.10.1'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*ipv4addr*used multiple times*')
                }
            } 
            It 'name' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF1.name = 'kiwi'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testRemoveWorksheet[1]
                    $testF2.name = 'kiwi'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*name*used multiple times*')
                }
            } 
            It 'ddns_hostname' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                    $testF1.ddns_hostname = 'kiwi'
                    $testF1

                    $testF2 = Copy-ObjectHC -Name $testRemoveWorksheet[1]
                    $testF2.ddns_hostname = 'kiwi'
                    $testF2
                } -ParameterFilter {
                    $WorksheetName -eq 'Remove'
                }

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailUsersParams) -and ($Message -like '*ddns_hostname*used multiple times*')
                }
            } 
        }
    }
}
Describe "duplicate address reservations in both 'FixedAddress' and 'Remove'" {
    Context 'are reported to the user as an error for' {
        It 'ipv4addr' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.ipv4addr = '10.10.10.1'
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.ipv4addr = '10.10.10.1'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*duplicate*')
            }
        } 
        It 'name' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.name = 'kiwi'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.name = 'kiwi'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*duplicate*')
            }
        } 
        It 'ddns_hostname' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.ddns_hostname = 'kiwi'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.ddns_hostname = 'kiwi'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*duplicate*')
            }
        } 
        It 'mac' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.mac = '00:00:00:00:00:01'
                $testF1.match_client = 'MAC_ADDRESS'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
                $testF1.mac = '00:00:00:00:00:01'
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'Remove'
            }

            .$testScript @testParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailUsersParams) -and ($Message -like '*duplicate*')
            }
        } 
    }
}
Describe "an address reservation in the worksheet 'FixedAddress' is" {
    Context 'created when it is not in IPAM and the client is' {
        BeforeAll {
            Mock Import-Excel
        }
        Context 'offline' {
            BeforeAll {
                Mock Test-Connection { $false }
            }
            It 'and overwrite is true' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.overwrite = $true
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke New-IpamFixedAddressHC -Times 1 -Exactly
            } 
            It 'and overwrite is false' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.overwrite = $false
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke New-IpamFixedAddressHC -Times 1 -Exactly
            } 
        }
        Context 'online' {
            BeforeAll {
                Mock Test-Connection { [PSCustomObject]@{
                        Address = '10.10.10.1'
                    } }
            }
            It 'and overwrite is true' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.overwrite = $true
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke New-IpamFixedAddressHC -Times 1 -Exactly
            } 
            It 'and is not created when overwrite is false' {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.overwrite = $false
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }

                .$testScript @testParams

                Should -Invoke New-IpamFixedAddressHC -Times 0 -Exactly
                $FixedAddressWorksheet.Status | Should -Be 'Error'
                $FixedAddressWorksheet.Error | Should -BeLike '*online*'
            } 
        }
    }
    Context 'updated when it is already in IPAM with an incorrect field' {
        BeforeAll {
            Mock Get-IpamFixedAddressHC -MockWith {
                $testF1 = Copy-ObjectHC $testFixedAddressWorksheet[0] |
                Select-Object -Property * -ExcludeProperty overwrite
                $testF1.Comment = 'Orig'
                $testF1
            }
        }
        Context 'and OverWrite is true and the client is' {
            BeforeAll {
                Mock Import-Excel {
                    $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                    $testF1.Comment = 'Diff'
                    $testF1.overwrite = $true
                    $testF1
                } -ParameterFilter {
                    $WorksheetName -eq 'FixedAddress'
                }
            }
            It 'offline' {
                Mock Test-Connection { $false }

                .$testScript @testParams

                $FixedAddressWorksheet.Status | Should -Be 'Updated'

                Should -Invoke Remove-IpamObjectHC -Times 1 -Exactly
                Should -Invoke New-IpamFixedAddressHC -Times 1 -Exactly
            } 
            It 'online' {
                Mock Test-Connection { [PSCustomObject]@{
                        Address = '10.10.10.1'
                    } }

                .$testScript @testParams

                $FixedAddressWorksheet.Status | Should -Be 'Updated'

                Should -Invoke Remove-IpamObjectHC -Times 1 -Exactly
                Should -Invoke New-IpamFixedAddressHC -Times 1 -Exactly
            } 
        }
        It 'and is not updated when OverWrite is false' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.overwrite = $false
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.overwrite = $null
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }

            .$testScript @testParams

            Should -Invoke Remove-IpamObjectHC -Times 0 -Exactly
            Should -Invoke New-IpamFixedAddressHC -Times 0 -Exactly
        } 
    }
    Context 'is marked as error when' {
        It 'multiple addresses have the same IP and/or mac' {
            Mock Import-Excel {
                $testF1 = Copy-ObjectHC -Name $testFixedAddressWorksheet[0]
                $testF1.overwrite = $false
                $testF1
            } -ParameterFilter {
                $WorksheetName -eq 'FixedAddress'
            }
            Mock Get-IpamFixedAddressHC -MockWith {
                [PSCustomObject]@{
                    name = $testFixedAddressWorksheet[0].name
                }
                [PSCustomObject]@{
                    name = $testFixedAddressWorksheet[0].name
                }
            }

            .$testScript @testParams

            $FixedAddressWorksheet.Status | Should -Be 'Error'
            $FixedAddressWorksheet.Error | Should -BeLike '*already known in IPAM*'

            Should -Invoke Remove-IpamObjectHC -Times 0 -Exactly
            Should -Invoke New-IpamFixedAddressHC -Times 0 -Exactly
        } 
    }
}
Describe "remove an address reservation when it's in the worksheet 'Remove'" {
    It "if it doesn't exist in IPAM nothing is removed and the Status is 'Not found'" {
        Mock Get-IpamFixedAddressHC
        Mock Import-Excel
        Mock Import-Excel {
            $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testF1.ipv4addr = '10.10.10.1'
            $testF1
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }

        .$testScript @testParams

        $RemoveWorksheet.Status | Should -Be 'Not found'
        Should -Invoke Remove-IpamObjectHC -Exactly 0
    } 
    It "if it does exist in IPAM remove it and set the Status to 'Removed'" {
        Mock Get-IpamFixedAddressHC {
            [PSCustomObject]@{
                name = $testFixedAddressWorksheet[0].name
            }
        }
        Mock Import-Excel
        Mock Import-Excel {
            $testF1 = Copy-ObjectHC -Name $testRemoveWorksheet[0]
            $testF1.ipv4addr = '10.10.10.1'
            $testF1
        } -ParameterFilter {
            $WorksheetName -eq 'Remove'
        }

        .$testScript @testParams

        $RemoveWorksheet.Status | Should -Be 'Removed'
        $RemoveWorksheet.Action | Should -BeLike '*removed*'
        Should -Invoke Remove-IpamObjectHC -Exactly 1
    } 
}
    
