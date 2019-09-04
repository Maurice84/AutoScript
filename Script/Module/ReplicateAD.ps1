Function Script-Module-ReplicateAD {
    # ------------------------------------------------------
    # Index all DCs and use REPADMIN to replicate to each DC
    # ------------------------------------------------------
    Script-Connect-Server -Command ([scriptblock]::Create(`
        "`$DomainControllers = Get-ADDomainController -Filter {HostName -ne `$PDC};
        If (`$DomainControllers) {
            ForEach (`$DC in `$DomainControllers) {
                REPADMIN /replicate `$DC.HostName `$PDC `$DC.DefaultPartition | Out-Null;
                REPADMIN /replicate `$PDC `$DC.HostName `$DC.DefaultPartition | Out-Null;
            }
        }")
    )
}