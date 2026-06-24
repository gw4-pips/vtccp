---
name: EthSystemDiscoverer SDK stub
description: Cognex DataMan SDK v25.4.1 does not expose EthSystemDiscoverer; discovery is stubbed
---

## Rule
`Cognex.DataMan.SDK.EthSystemDiscoverer` does NOT exist in Cognex DataMan SDK v25.4.1
(`Cognex.DataMan.SDK.PC.dll` from DataMan Software v25.4.1).
`EthSystemConnector` and `DataManSystem` ARE present and working.
`NetworkDiscoverer.DiscoverAsync` is currently stubbed to return an empty list.

**Why:** The class name was written speculatively based on older SDK docs. The compile
confirmed it doesn't exist in v25.4.1.

**How to apply:**
To restore SDK-based network discovery, confirm the correct class name by:
1. Opening ILSpy / dotPeek on `Cognex.DataMan.SDK.PC.dll`
2. Searching for any class matching `*Discover*` or `*Discovery*` in the Cognex.DataMan.SDK namespace
3. Checking Cognex SDK release notes for v25.x API changes
Once confirmed, replace the stub body in `vtccp/DeviceInterface/NetworkDiscoverer.cs`.
Manual IP entry via ⊕ Import remains available in the meantime.
