Alright, I've pulled together a solid picture of what's going on. This is a well-known class of problems with IPv6 in large vSphere environments. Here's the breakdown:

Root Causes: Why IPv6 Broke Everything
1. IPv6 Multicast Flooding on the Physical Switches
This is the most likely primary culprit. When you enable IPv6 on ESXi hosts, each VMkernel adapter immediately generates link-local addresses and begins participating in IPv6 Neighbor Discovery Protocol (NDP). NDP relies heavily on multicast — Router Solicitations (RS), Router Advertisements (RA), Neighbor Solicitations (NS), and Duplicate Address Detection (DAD) all use multicast groups (ff02::1, ff02::2, ff02::1:ffxx:xxxx). If your physical switches aren't configured with MLD snooping (the IPv6 equivalent of IGMP snooping), all this multicast traffic gets flooded to every port in the VLAN. At 4,000+ hosts, that's a massive amount of multicast being treated as broadcast, which will absolutely cause MAC flapping on the physical switches and degrade the network.
2. SLAAC Auto-Addressing Chaos
When IPv6 is enabled, each VMkernel interface auto-generates a link-local address via SLAAC. If there are routers on those VLANs sending Router Advertisements, the hosts will also auto-configure global IPv6 addresses. VMware's HA best practices explicitly state that only one IPv6 address can be assigned to a given network interface used by a vSphere HA cluster, because assigning multiple IP addresses increases the number of heartbeat messages sent with no corresponding benefit VMware. SLAAC can easily result in multiple addresses per interface (link-local + global + temporary privacy addresses), which directly breaks this guidance.
3. HA Heartbeat Confusion in Dual-Stack Mode
HA heartbeats travel over management VMkernel networks. When IPv6 is added, HA tries to use both IPv4 and IPv6 for heartbeat communication. If the IPv6 path is unreliable (due to the multicast flooding or routing issues above), hosts intermittently fail to receive heartbeats, causing false isolation events and HA instability. The vmkNICs used by vSphere HA and by other features must be on different subnets VMware to ensure proper flow separation — and adding IPv6 subnets without planning this out creates overlap.
4. vMotion Address Mismatch / Routing Issues
vMotion requires that source and destination hosts can reach each other on the vMotion VMkernel interface. If one host has picked up an IPv6 address via SLAAC and another hasn't (or got a different prefix), vMotion will fail because it can't establish the connection. There's also a known constraint that you cannot vMotion from a vCenter Server with an IPv4 address to a host registered with an IPv6 address VMware — mixing addressing modes causes failures.

Resolution Plan for Re-enabling IPv6
Here's the approach I'd recommend for rolling this out properly at your scale:
Phase 1: Physical Network Prep (do this BEFORE touching ESXi)

Enable MLD snooping on all physical switches carrying ESXi traffic VLANs. This is the single most important step — it prevents IPv6 multicast from being flooded as broadcast.
Configure RA Guard on switch ports facing ESXi hosts to prevent rogue Router Advertisements. Only your designated IPv6 routers should be sending RAs.
Suppress Router Advertisements on VLANs/interfaces that don't need them (vMotion, vSAN, storage VLANs). RAs should only be active on management VLANs where you actually want SLAAC or DHCPv6.
Work with your network team to confirm IPv6 routing is in place and that the VLANs have proper prefix assignments.

Phase 2: VMkernel IPv6 Configuration

Use static IPv6 addresses (or DHCPv6 with reservations) on VMkernel adapters rather than relying on SLAAC. This gives you predictable addressing and avoids the multiple-address problem.
Assign only one IPv6 address per VMkernel interface per VMware's HA best practices.
If you're using the vMotion TCP/IP stack, configure the IPv6 settings on it with a dedicated gateway.
Set the HA isolation address (das.isolationaddress) to include an IPv6 gateway you control, so HA isolation detection works correctly in dual-stack.

Phase 3: Phased Rollout

Start with a single non-production cluster. Enable IPv6 on the hosts, reboot, and validate HA heartbeats, vMotion, and network stability over 48-72 hours.
Monitor switch logs for MAC flapping and multicast flooding.
Use vmkping -6 between hosts to verify IPv6 connectivity on each VMkernel interface.
Gradually expand to production clusters once validated.

Phase 4: Ongoing Monitoring

Monitor vmkernel.log for link flapping events — ESXi will auto-disable a NIC if it flaps more than 60 times per minute (controlled by Net.LinkFlappingThreshold).
Watch for HA reconfiguration events or false isolation alarms.
Keep MLD snooping querier active on your switches.


The TL;DR is that the network team and the VMware team need to coordinate on this one. The ESXi side is actually pretty straightforward, but if the physical switches aren't ready for IPv6 multicast, everything falls apart. MLD snooping is the linchpin — without it, enabling IPv6 on a large vSphere environment is basically a self-inflicted broadcast storm.