Function Messaging-Exchange-Create-Customer-3-Confirmation {
    # =========
    # Execution
    # =========
    Invoke-Expression -Command ( `
		Script-Function-Confirmation `
			-Name ($MyInvocation.MyCommand).Name `
    )
}