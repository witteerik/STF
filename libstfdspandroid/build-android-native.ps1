param(
	[string]$AndroidSdkRoot,
	[string]$AndroidNdkRoot,
	[string]$Configuration = "Release",
	[string]$AndroidPlatform = "android-31",
	[string[]]$Abis = @("armeabi-v7a", "arm64-v8a", "x86", "x86_64"),
	[switch]$Clean,
	[switch]$SkipPublish
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Get-NormalizedPathCandidates {
	param([string]$PathValue)

	if ([string]::IsNullOrWhiteSpace($PathValue)) {
		return @()
	}

	$candidates = New-Object System.Collections.Generic.List[string]
	$trimmed = $PathValue.Trim().Trim('"')
	$candidates.Add($trimmed)

	if ($trimmed.Contains('"')) {
		$quotedPrefix = $trimmed.Split('"')[0].Trim()
		if (-not [string]::IsNullOrWhiteSpace($quotedPrefix)) {
			$candidates.Add($quotedPrefix)
		}
	}

	$decodedCandidates = New-Object System.Collections.Generic.List[string]
	foreach ($candidate in $candidates) {
		$decodedCandidates.Add($candidate)
		$decoded = [System.Uri]::UnescapeDataString($candidate)
		if ($decoded -ne $candidate) {
			$decodedCandidates.Add($decoded)
		}
	}

	return $decodedCandidates | Select-Object -Unique
}

function Try-ResolveExistingPath {
	param([string]$PathValue)

	if ([string]::IsNullOrWhiteSpace($PathValue)) {
		return $null
	}

	try {
		if (Test-Path -LiteralPath $PathValue) {
			return (Resolve-Path -LiteralPath $PathValue).Path
		}
	}
	catch {
		return $null
	}

	return $null
}

function Get-AndroidSdkCandidatesFromNdkPath {
	param([string]$NdkPath)

	$candidates = New-Object System.Collections.Generic.List[string]
	foreach ($normalizedCandidate in (Get-NormalizedPathCandidates -PathValue $NdkPath)) {
		$trimmedCandidate = $normalizedCandidate.TrimEnd('\')
		if ($trimmedCandidate -like '*\ndk-bundle') {
			$candidates.Add((Split-Path -Parent $trimmedCandidate))
			continue
		}

		$parentPath = Split-Path -Parent $trimmedCandidate
		if (-not [string]::IsNullOrWhiteSpace($parentPath) -and (Split-Path -Leaf $parentPath) -ieq 'ndk') {
			$candidates.Add((Split-Path -Parent $parentPath))
			continue
		}

		$resolvedNdkPath = Try-ResolveExistingPath -PathValue $normalizedCandidate
		if ($null -eq $resolvedNdkPath) {
			continue
		}

		$leafName = Split-Path -Leaf $resolvedNdkPath
		if ($leafName -ieq "ndk-bundle") {
			$candidates.Add((Split-Path -Parent $resolvedNdkPath))
			continue
		}

		$resolvedParentPath = Split-Path -Parent $resolvedNdkPath
		if ((Split-Path -Leaf $resolvedParentPath) -ieq "ndk") {
			$candidates.Add((Split-Path -Parent $resolvedParentPath))
		}
	}

	return $candidates | Select-Object -Unique
}

function Resolve-AndroidSdkRoot {
	param(
		[string]$PreferredPath,
		[string]$PreferredNdkPath
	)

	$candidates = @(
		$PreferredPath,
		$env:ANDROID_SDK_ROOT,
		$env:ANDROID_HOME,
		(Join-Path $env:LOCALAPPDATA "Android\Sdk")
	) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

	$ndkCandidates = @(
		$PreferredNdkPath,
		$env:ANDROID_NDK_ROOT,
		$env:ANDROID_NDK_HOME
	) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

	foreach ($ndkCandidate in $ndkCandidates) {
		$candidates += Get-AndroidSdkCandidatesFromNdkPath -NdkPath $ndkCandidate
	}

	foreach ($candidate in $candidates) {
		foreach ($normalizedCandidate in (Get-NormalizedPathCandidates -PathValue $candidate)) {
			$resolvedPath = Try-ResolveExistingPath -PathValue $normalizedCandidate
			if ($null -ne $resolvedPath) {
				return $resolvedPath
			}
		}
	}

	throw "Android SDK root could not be resolved. Pass -AndroidSdkRoot or set ANDROID_SDK_ROOT/ANDROID_HOME."
}

function Resolve-AndroidNdkRoot {
	param(
		[string]$PreferredPath,
		[string]$SdkRoot
	)

	$candidates = New-Object System.Collections.Generic.List[string]
	if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
		$candidates.Add($PreferredPath)
	}

	if (-not [string]::IsNullOrWhiteSpace($env:ANDROID_NDK_ROOT)) {
		$candidates.Add($env:ANDROID_NDK_ROOT)
	}

	if (-not [string]::IsNullOrWhiteSpace($env:ANDROID_NDK_HOME)) {
		$candidates.Add($env:ANDROID_NDK_HOME)
	}

	$sdkNdkBundle = Join-Path $SdkRoot "ndk-bundle"
	if (Test-Path $sdkNdkBundle) {
		$candidates.Add($sdkNdkBundle)
	}

	$sdkNdkRoot = Join-Path $SdkRoot "ndk"
	if (Test-Path $sdkNdkRoot) {
		$latestNdk = Get-ChildItem -Path $sdkNdkRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
		if ($null -ne $latestNdk) {
			$candidates.Add($latestNdk.FullName)
		}
	}

	foreach ($candidate in $candidates) {
		foreach ($normalizedCandidate in (Get-NormalizedPathCandidates -PathValue $candidate)) {
			$resolvedPath = Try-ResolveExistingPath -PathValue $normalizedCandidate
			if ($null -ne $resolvedPath) {
				return $resolvedPath
			}
		}
	}

	throw "Android NDK root could not be resolved. Pass -AndroidNdkRoot or set ANDROID_NDK_ROOT/ANDROID_NDK_HOME."
}

function Assert-CommandAvailable {
	param([string]$CommandName)

	if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) {
		throw "Required command '$CommandName' was not found on PATH."
	}
}

function Resolve-CommandPath {
	param(
		[string]$CommandName,
		[string[]]$CandidatePaths = @()
	)

	foreach ($candidate in $CandidatePaths) {
		foreach ($normalizedCandidate in (Get-NormalizedPathCandidates -PathValue $candidate)) {
			$resolvedPath = Try-ResolveExistingPath -PathValue $normalizedCandidate
			if ($null -ne $resolvedPath) {
				return $resolvedPath
			}
		}
	}

	$command = Get-Command $CommandName -ErrorAction SilentlyContinue
	if ($null -ne $command) {
		return $command.Source
	}

	throw "Required command '$CommandName' was not found on PATH and no fallback path was valid."
}

function Resolve-AndroidSdkCMakeRoot {
	param([string]$SdkRoot)

	$sdkCMakeParent = Join-Path $SdkRoot "cmake"
	if (-not (Test-Path $sdkCMakeParent)) {
		return $null
	}

	$latestCMake = Get-ChildItem -Path $sdkCMakeParent -Directory | Sort-Object Name -Descending | Select-Object -First 1
	if ($null -eq $latestCMake) {
		return $null
	}

	return $latestCMake.FullName
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$stfResearchSuiteRoot = Join-Path $repoRoot "StfResearchSuite"
$nativeLibOutputRoot = Join-Path $stfResearchSuiteRoot "Platforms\Android\native-libs"
$cmakeSourceDir = $scriptRoot
$outRoot = Join-Path $scriptRoot "out"

$resolvedSdkRoot = Resolve-AndroidSdkRoot -PreferredPath $AndroidSdkRoot -PreferredNdkPath $AndroidNdkRoot
$resolvedNdkRoot = Resolve-AndroidNdkRoot -PreferredPath $AndroidNdkRoot -SdkRoot $resolvedSdkRoot
$resolvedSdkCMakeRoot = Resolve-AndroidSdkCMakeRoot -SdkRoot $resolvedSdkRoot
$toolchainFile = Join-Path $resolvedNdkRoot "build\cmake\android.toolchain.cmake"
$cmakeExecutable = Resolve-CommandPath -CommandName "cmake" -CandidatePaths @(
	(Join-Path $resolvedSdkCMakeRoot "cmake.exe"),
	(Join-Path $resolvedSdkCMakeRoot "bin\cmake.exe")
)
$ninjaExecutable = Resolve-CommandPath -CommandName "ninja" -CandidatePaths @(
	(Join-Path $resolvedSdkCMakeRoot "ninja.exe"),
	(Join-Path $resolvedSdkCMakeRoot "bin\ninja.exe")
)

if (-not (Test-Path $toolchainFile)) {
	throw "Android NDK toolchain file was not found at '$toolchainFile'."
}

Write-Host "Android SDK root: $resolvedSdkRoot"
Write-Host "Android NDK root: $resolvedNdkRoot"
Write-Host "Android SDK CMake root: $resolvedSdkCMakeRoot"
Write-Host "CMake executable: $cmakeExecutable"
Write-Host "Ninja executable: $ninjaExecutable"
Write-Host "Configuration: $Configuration"
Write-Host "Android platform: $AndroidPlatform"
Write-Host "ABIs: $($Abis -join ', ')"

foreach ($abi in $Abis) {
	$abiBuildDir = Join-Path $outRoot $abi
	$abiPublishDir = Join-Path $nativeLibOutputRoot $abi

	if ($Clean -and (Test-Path $abiBuildDir)) {
		Remove-Item -Path $abiBuildDir -Recurse -Force
	}

	if (-not (Test-Path $abiBuildDir)) {
		New-Item -ItemType Directory -Path $abiBuildDir | Out-Null
	}

	Write-Host "Configuring ABI '$abi'..."
	& $cmakeExecutable -S $cmakeSourceDir -B $abiBuildDir -G Ninja `
		"-DCMAKE_MAKE_PROGRAM=$ninjaExecutable" `
		"-DCMAKE_TOOLCHAIN_FILE=$toolchainFile" `
		"-DANDROID_ABI=$abi" `
		"-DANDROID_PLATFORM=$AndroidPlatform" `
		"-DCMAKE_BUILD_TYPE=$Configuration"

	if ($LASTEXITCODE -ne 0) {
		throw "CMake configure failed for ABI '$abi'."
	}

	Write-Host "Building ABI '$abi'..."
	& $cmakeExecutable --build $abiBuildDir --config $Configuration

	if ($LASTEXITCODE -ne 0) {
		throw "CMake build failed for ABI '$abi'."
	}

	$builtLibraryPath = Join-Path $abiBuildDir "liblibostfdspandroid.so"
	if (-not (Test-Path $builtLibraryPath)) {
		throw "Expected built library was not found at '$builtLibraryPath'."
	}

	if (-not $SkipPublish) {
		if (-not (Test-Path $abiPublishDir)) {
			New-Item -ItemType Directory -Path $abiPublishDir -Force | Out-Null
		}

		Copy-Item -Path $builtLibraryPath -Destination (Join-Path $abiPublishDir "liblibostfdspandroid.so") -Force
		Write-Host "Published '$abi' to '$abiPublishDir'."
	}
}

Write-Host "Android native build completed successfully."
if ($SkipPublish) {
	Write-Host "Publish step was skipped."
}
