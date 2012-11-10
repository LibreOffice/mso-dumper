#!/usr/bin/env python

# see 2.6.2 of the spec
parMap = {
    0x4600: "sprmPIstd",
    0xC601: "sprmPIstdPermute",
    0x2602: "sprmPIncLvl",
    0x2403: "sprmPJc80",
    0x2405: "sprmPFKeep",
    0x2406: "sprmPFKeepFollow",
    0x2407: "sprmPFPageBreakBefore",
    0x260A: "sprmPIlvl",
    0x460B: "sprmPIlfo",
    0x240C: "sprmPFNoLineNumb",
    0xC60D: "sprmPChgTabsPapx",
    0x840E: "sprmPDxaRight80",
    0x840F: "sprmPDxaLeft80",
    0x4610: "sprmPNest80",
    0x8411: "sprmPDxaLeft180",
    0x6412: "sprmPDyaLine",
    0xA413: "sprmPDyaBefore",
    0xA414: "sprmPDyaAfter",
    0xC615: "sprmPChgTabs",
    0x2416: "sprmPFInTable",
    0x2417: "sprmPFTtp",
    0x8418: "sprmPDxaAbs",
    0x8419: "sprmPDyaAbs",
    0x841A: "sprmPDxaWidth",
    0x261B: "sprmPPc",
    0x2423: "sprmPWr",
    0x6424: "sprmPBrcTop80",
    0x6425: "sprmPBrcLeft80",
    0x6426: "sprmPBrcBottom80",
    0x6427: "sprmPBrcRight80",
    0x6428: "sprmPBrcBetween80",
    0x6629: "sprmPBrcBar80",
    0x242A: "sprmPFNoAutoHyph",
    0x442B: "sprmPWHeightAbs",
    0x442C: "sprmPDcs",
    0x442D: "sprmPShd80",
    0x842E: "sprmPDyaFromText",
    0x842F: "sprmPDxaFromText",
    0x2430: "sprmPFLocked",
    0x2431: "sprmPFWidowControl",
    0x2433: "sprmPFKinsoku",
    0x2434: "sprmPFWordWrap",
    0x2435: "sprmPFOverflowPunct",
    0x2436: "sprmPFTopLinePunct",
    0x2437: "sprmPFAutoSpaceDE",
    0x2438: "sprmPFAutoSpaceDN",
    0x4439: "sprmPWAlignFont",
    0x443A: "sprmPFrameTextFlow",
    0x2640: "sprmPOutLvl",
    0x2441: "sprmPFBiDi",
    0x2443: "sprmPFNumRMIns",
    0xC645: "sprmPNumRM",
    0x6646: "sprmPHugePapx",
    0x2447: "sprmPFUsePgsuSettings",
    0x2448: "sprmPFAdjustRight",
    0x6649: "sprmPItap",
    0x664A: "sprmPDtap",
    0x244B: "sprmPFInnerTableCell",
    0x244C: "sprmPFInnerTtp",
    0xC64D: "sprmPShd",
    0xC64E: "sprmPBrcTop",
    0xC64F: "sprmPBrcLeft",
    0xC650: "sprmPBrcBottom",
    0xC651: "sprmPBrcRight",
    0xC652: "sprmPBrcBetween",
    0xC653: "sprmPBrcBar",
    0x4455: "sprmPDxcRight",
    0x4456: "sprmPDxcLeft",
    0x4457: "sprmPDxcLeft1",
    0x4458: "sprmPDylBefore",
    0x4459: "sprmPDylAfter",
    0x245A: "sprmPFOpenTch",
    0x245B: "sprmPFDyaBeforeAuto",
    0x245C: "sprmPFDyaAfterAuto",
    0x845D: "sprmPDxaRight",
    0x845E: "sprmPDxaLeft",
    0x465F: "sprmPNest",
    0x8460: "sprmPDxaLeft1",
    0x2461: "sprmPJc",
    0x2462: "sprmPFNoAllowOverlap",
    0x2664: "sprmPWall",
    0x6465: "sprmPIpgp",
    0xC666: "sprmPCnf",
    0x6467: "sprmPRsid",
    0xC669: "sprmPIstdListPermute",
    0x646B: "sprmPTableProps",
    0xC66C: "sprmPTIstdInfo",
    0x246D: "sprmPFContextualSpacing",
    0xC66F: "sprmPPropRMark",
    0x2470: "sprmPFMirrorIndents",
    0x2471: "sprmPTtwo",
        }

# see 2.6.1 of the spec
chrMap = {
    0x0800: "sprmCFRMarkDel",
    0x0801: "sprmCFRMarkIns",
    0x0802: "sprmCFFldVanish",
    0x6A03: "sprmCPicLocation",
    0x4804: "sprmCIbstRMark",
    0x6805: "sprmCDttmRMark",
    0x0806: "sprmCFData",
    0x4807: "sprmCIdslRMark",
    0x6A09: "sprmCSymbol",
    0x080A: "sprmCFOle2",
    0x2A0C: "sprmCHighlight",
    0x0811: "sprmCFWebHidden",
    0x6815: "sprmCRsidProp",
    0x6816: "sprmCRsidText",
    0x6817: "sprmCRsidRMDel",
    0x0818: "sprmCFSpecVanish",
    0xC81A: "sprmCFMathPr",
    0x4A30: "sprmCIstd",
    0xCA31: "sprmCIstdPermute",
    0x2A33: "sprmCPlain",
    0x2A34: "sprmCKcd",
    0x0835: "sprmCFBold",
    0x0836: "sprmCFItalic",
    0x0837: "sprmCFStrike",
    0x0838: "sprmCFOutline",
    0x0839: "sprmCFShadow",
    0x083A: "sprmCFSmallCaps",
    0x083B: "sprmCFCaps",
    0x083C: "sprmCFVanish",
    0x2A3E: "sprmCKul",
    0x8840: "sprmCDxaSpace",
    0x2A42: "sprmCIco",
    0x4A43: "sprmCHps",
    0x4845: "sprmCHpsPos",
    0xCA47: "sprmCMajority",
    0x2A48: "sprmCIss",
    0x484B: "sprmCHpsKern",
    0x484E: "sprmCHresi",
    0x4A4F: "sprmCRgFtc0",
    0x4A50: "sprmCRgFtc1",
    0x4A51: "sprmCRgFtc2",
    0x4852: "sprmCCharScale",
    0x2A53: "sprmCFDStrike",
    0x0854: "sprmCFImprint",
    0x0855: "sprmCFSpec",
    0x0856: "sprmCFObj",
    0xCA57: "sprmCPropRMark90",
    0x0858: "sprmCFEmboss",
    0x2859: "sprmCSfxText",
    0x085A: "sprmCFBiDi",
    0x085C: "sprmCFBoldBi",
    0x085D: "sprmCFItalicBi",
    0x4A5E: "sprmCFtcBi",
    0x485F: "sprmCLidBi",
    0x4A60: "sprmCIcoBi",
    0x4A61: "sprmCHpsBi",
    0xCA62: "sprmCDispFldRMark",
    0x4863: "sprmCIbstRMarkDel",
    0x6864: "sprmCDttmRMarkDel",
    0x6865: "sprmCBrc80",
    0x4866: "sprmCShd80",
    0x4867: "sprmCIdslRMarkDel",
    0x0868: "sprmCFUsePgsuSettings",
    0x486D: "sprmCRgLid0_80",
    0x486E: "sprmCRgLid1_80",
    0x286F: "sprmCIdctHint",
    0x6870: "sprmCCv",
    0xCA71: "sprmCShd",
    0xCA72: "sprmCBrc",
    0x4873: "sprmCRgLid0",
    0x4874: "sprmCRgLid1",
    0x0875: "sprmCFNoProof",
    0xCA76: "sprmCFitText",
    0x6877: "sprmCCvUl",
    0xCA78: "sprmCFELayout",
    0x2879: "sprmCLbcCRJ",
    0x0882: "sprmCFComplexScripts",
    0x2A83: "sprmCWall",
    0xCA85: "sprmCCnf",
    0x2A86: "sprmCNeedFontFixup",
    0x6887: "sprmCPbiIBullet",
    0x4888: "sprmCPbiGrf",
    0xCA89: "sprmCPropRMark",
    0x2A90: "sprmCFSdtVanish",
        }

# vim:set filetype=python shiftwidth=4 softtabstop=4 expandtab: