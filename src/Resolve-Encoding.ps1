<#
	.SYNOPSIS
		テキストエンコーディングを判別します。
	
	.DESCRIPTION
		ファイルもしくはバイトストリームからテキストエンコーディングを判別します。
	
	.PARAMETER InputObject
		テキストを表すバイト配列です。
	
	.PARAMETER Path
		テキストファイル名です。
		ワイルドカードが使用可能です。
	
	.PARAMETER LiteralPath
		テキストファイル名です。
		ワイルドカードは使えません。
	
	.PARAMETER MaxInputCount
		先頭から MaxInputCount で示されるバイト数だけ判別対象とします。
		規定値は2048です。
	
	.INPUTS
		System.Byte[]
		パイプを利用して InputObject を渡すことができます。
	
	.OUTPUTS
		System.Text.Encoding
	
	.EXAMPLE
		Resolve-Encoding .\euc.txt
		.\euc.txt のエンコーディングを取得します。
	
	.EXAMPLE
		Get-Content -Encoding Byte .\euc.txt | Resolve-Encoding
		.\euc.txt のエンコーディングを取得します。
	
	.EXAMPLE
		Get-ChildItem *.txt | Foreach-Object { $_ | Select-Object @{Name='ファイル名';Expression={$_.Name}},@{Name='エンコーディング';Expression={(Resolve-Encoding $_).EncodingName}} }
		現在位置の *.txt ファイルのエンコーディングを表にして出力します。
	
	.LINK
		http://flamework.net
#>
[CmdletBinding(DefaultParameterSetName = 'Path')]
Param (
	[Parameter(Mandatory = $True, Position = 0, ParameterSetName = 'InputObject', ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
	[Byte[]]$InputObject,
	
	[Parameter(Mandatory = $True, Position = 0, ParameterSetName = 'Path', ValueFromPipelineByPropertyName = $True)]
	[String[]]$Path,
	
	[Parameter(Mandatory = $True, Position = 0, ParameterSetName = 'LiteralPath', ValueFromPipelineByPropertyName = $True)]
	[String[]]$LiteralPath,
	
	[Int]$MaxInputCount = 2048
)
Begin {
	Write-Debug "Begin"
	function createAutomaton([String]$Name, [Text.Encoding]$Encoding, [ScriptBlock]$Work) {
		New-Object Object | Select-Object `
			@{Name = 'Name'; Expression = {$Name}},
			@{Name = 'Status'; Expression = {1}},
			@{Name = 'Encoding'; Expression = {$Encoding}} |
				Add-Member ScriptMethod Work $Work -PassThru |
				Add-Member ScriptMethod Succeed {$this.Status = 0} -PassThru |
				Add-Member ScriptMethod Fail {$this.Status = -1} -PassThru |
				Add-Member ScriptMethod IsSuccess {$this.Status -eq 0} -PassThru |
				Add-Member ScriptMethod IsFailed {$this.Status -lt 0} -PassThru
	}
	function initialize {
		$script:automatons = New-Object Collections.ArrayList
		$script:automatons.AddRange(@(
			(createAutomaton UTF32 ([Text.Encoding]::UTF32) {
				Param([Int]$c)
				switch ($this.Status) {
					1 {
						if ($c -eq 0xFF) { $this.Status = 2 }
						else { $this.Fail() }
					}
					2 {
						if ($c -eq 0xFE) { $this.Status = 3 }
						else { $this.Fail() }
					}
					3 {
						if ($c -eq 0) { $this.Status = 4 }
						else { $this.Fail() }
					}
					4 {
						if ($c -eq 0) { $this.Succeed() }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton UTF32Be ([Text.Encoding]::GetEncoding('UTF-32BE')) {
				Param([Int]$c)
				switch ($this.Status) {
					1 {
						if ($c -eq 0) { $this.Status = 2 }
						else { $this.Fail() }
					}
					2 {
						if ($c -eq 0) { $this.Status = 3 }
						else { $this.Fail() }
					}
					3 {
						if ($c -eq 0xFE) { $this.Status = 4 }
						else { $this.Fail() }
					}
					4 {
						if ($c -eq 0xFF) { $this.Succeed() }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton UTF16 ([Text.Encoding]::Unicode) {
				Param([Int]$c)
				switch ($this.Status) {
					1 {
						if ($c -eq 0xFF) { $this.Status = 2 }
						else { $this.Fail() }
					}
					2 {
						if ($c -eq 0xFE) { $this.Status = 3 }
						else { $this.Fail() }
					}
					3 {
						if ($c -eq 0) { $this.Status = 4 }
						else { $this.Succeed() }
					}
					4 {
						if ($c -eq 0) { $this.Fail() }
						else { $this.Succeed() }
					}
				}
			}),
			(createAutomaton UTF16Be ([Text.Encoding]::BigEndianUnicode) {
				Param([Int]$c)
				switch ($this.Status) {
					1 {
						if ($c -eq 0xFE) { $this.Status = 2 }
						else { $this.Fail() }
					}
					2 {
						if ($c -eq 0xFF) { $this.Succeed() }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton ASCII ([Text.Encoding]::ASCII) {
				Param([Int]$c)
				if ($c -eq -1) {
					$this.Succeed()
				} elseif ($c -gt 0x7F) {
					$this.Fail()
				}
			}),
			(createAutomaton JIS ([Text.Encoding]::GetEncoding('iso-2022-jp')) {
				Param([Int]$c)
				switch ($this.Status) {
					1 {
						if ($c -eq 0x1B) { $this.Status = 2 }
						elseif ($c -gt 0x7F) { $this.Fail() }
					}
					2 {
						switch ($c) {
							0x24 { $this.Status = 3 }
							0x28 { $this.Status = 4 }
							default { $this.Fail() }
						}
					}
					3 {
						if ($c -eq 0x40 -or $c -eq 0x42) { $this.Succeed() }
						elseif ($c -eq 0x28) { $this.Status = 5 }
						else { $this.Fail() }
					}
					4 {
						if ((0x42, 0x48, 0x49, 0x4A) -contains $c) { $this.Succeed() }
						else { $this.Fail() }
					}
					5 {
						if ((0x44, 0x4F, 0x50) -contains $c) { $this.Succeed() }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton UTF8 ([Text.Encoding]::UTF8) {
				Param([Int]$c)
				if ($c -eq -1) {
					$this.Succeed()
				} else {
					if ($this.Status -eq 1) {
						if ($c -eq 0xEF) { $this.Status = 10}
						elseif ($c -gt 0xFD) { $this.Fail() }
						elseif ($c -ge 0xFC) { $this.Status = 2 }
						elseif ($c -ge 0xF8) { $this.Status = 3 }
						elseif ($c -ge 0xF0) { $this.Status = 4 }
						elseif ($c -ge 0xE0) { $this.Status = 5 }
						elseif ($c -ge 0xC2) { $this.Status = 6 }
						elseif ($c -gt 0x7F) { $this.Fail() }
					} elseif ((2, 3, 4, 5) -contains $this.Status) {
						if ($c -ge 0x80 -and $c -le 0xBF) { $this.Status++ }
						else { $this.Fail() }
					} elseif ($this.Status -eq 6) {
						if ($c -ge 0x80 -and $c -le 0xBF) { $this.Status = 1 }
						else { $this.Fail() }
					} elseif ($this.Status -eq 10) {
						if ($c -eq 0xBB) { $this.Status++ }
						else { $this.Fail() }
					} elseif ($this.Status -eq 11) {
						if ($c -eq 0xBF) { $this.Succeed() }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton Euc ([Text.Encoding]::GetEncoding('euc-jp')) {
				Param([Int]$c)
				if ($c -eq -1) {
					$this.Succeed()
					return
				}
				switch ($this.Status) {
					1 {
						if ($c -ge 0xA1 -and $c -le 0xFE) { $this.Status = 2 }
						elseif ($c -eq 0x8E) { $this.Status = 3 }
						elseif ($c -eq 0x8F) { $this.Status = 4 }
						elseif ($c -eq 0xA4) { $this.Status = 5 }
						elseif ($c -eq 0xA5) { $this.Status = 6 }
						elseif ($c -eq 0xEF) { $this.Status = 8 }
						elseif ($c -gt 0x7F) { $this.Fail() }
					}
					2 {
						if ($c -ge 0xA1 -and $c -le 0xFE) { $this.Status = 1 }
						else { $this.Fail() }
					}
					3 {
						if ($c -ge 0xA1 -and $c -le 0xDF) { $this.Status = 1 }
						else { $this.Fail() }
					}
					4 {
						if ($c -ge 0xA1 -and $c -le 0xF3) { $this.Status = 1 }
						else { $this.Fail() }
					}
					5 {
						if ($c -ge 0xA1 -and $c -le 0xF6) { $this.Status = 1 }
						else { $this.Fail() }
					}
					6 {
						if ($c -ge 0xA1 -and $c -le 0xFE) { $this.Status = 7 }
						else { $this.Fail() }
					}
					7 {
						if ($c -ge 0xA1 -and $c -le 0xFE) { $this.Status = 1 }
						else { $this.Fail() }
					}
					8 {
						if ($c -eq 0xBB) { $this.Status = 9 }
						else { $this.Fail() }
					}
					9 {
						if ($c -eq 0xBF) { $this.Success }
						else { $this.Fail() }
					}
				}
			}),
			(createAutomaton Sjis ([Text.Encoding]::GetEncoding('shift-jis')) {
				Param([Int]$c)
				if ($c -eq -1) {
					$this.Succeed()
					return
				}
				switch ($this.Status) {
					1 {
						if (($c -ge 0x81 -and $c -le 0x9F) -or ($c -ge 0xE0)) { $this.Status++ }
						elseif (($c -gt 0x7F -and $c -lt 0xA1) -or ($c -gt 0xDF)) { $this.Fail() }
					}
					2 {
						if (($c -ge 0x40 -and $c -le 0x7E) -or ($c -ge 0x80 -and $c -le 0xFC)) { $this.Status-- }
						else { $this.Fail() }
					}
				}
			})
		))
		$script:inputCount = 0;
	}
	function work {
		Param([Int]$c)
		if ($script:inputCount -lt $MaxInputCount) {
			$script:inputCount++
			$script:automatons.ToArray() | Foreach-Object {
				if ($script:inputCount -lt $MaxInputCount) {
					$_.Work($c)
					if ($_.IsSuccess()) {
						$script:automatons.Clear()
						$script:automatons.AddRange(@($_))
						$script:inputCount = $MaxInputCount
					} elseif ($_.IsFailed()) {
						Write-Debug ("{0} : {1} 除外" -f $script:inputCount, $_.Name)
						$script:automatons.Remove($_)
					}
				}
			}
			$False
		} else {
			$True
		}
	}
	function finalize {
		[void](work -1)
		if ($script:automatons.Count -ne 0) {
			$script:automatons[0].Encoding
		}
	}
}
Process {
	Write-Debug "Process"
	switch ($pscmdlet.ParameterSetName) {
		'InputObject' {
			Write-Debug 'Data を処理しています。'
			if ($initialized -ne $true) {
				initialize
				$initialized = $true
			}
			foreach ($o in $InputObject) {
				if ((work $o)) { break }
			}
		}
		'Path' {
			Write-Debug 'Path を処理しています。'
			Get-Item $Path | Foreach-Object {
				Write-Debug $_
				initialize
				Get-Content $_ -Encoding Byte -TotalCount $MaxInputCount | Foreach-Object {
					[void](work ([Byte]$_))
				}
				finalize
			}
		}
		'LiteralPath' {
			Write-Debug 'LiteralPath を処理しています。'
			Get-Item -LiteralPath $LiteralPath | Foreach-Object {
				Write-Debug $_
				initialize
				Get-Content -LiteralPath $_ -Encoding Byte -TotalCount $MaxInputCount | Foreach-Object {
					[void](work ([Byte]$_))
				}
				finalize
			}
		}
	}
}
End {
	Write-Debug "End"
	if ($pscmdlet.ParameterSetName -eq 'InputObject') { finalize }
}
