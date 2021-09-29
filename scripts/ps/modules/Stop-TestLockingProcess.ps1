function Stop-TestLockingProcess($Path) {
	if (Test-Path -Path $Path) { Stop-LockingProcess -Path $Path	}
}