from pydantic import BaseModel, Field


class KbArticle(BaseModel):
    articleName: str = ""
    articleUrl: str = ""
    downloadName: str = ""
    downloadUrl: str = ""
    supercedence: str = ""
    rebootRequired: str = ""
    ordinal: int = 0
    fixedBuildNumber: str = ""


class AffectedProduct(BaseModel):
    ap_id: str = Field(alias="id")
    releaseDate: str = ""
    releaseNumber: str = ""
    product: str = ""
    productId: int = 0
    productFamily: str = ""
    productFamilyId: int = 0
    platform: str = ""
    platformId: int = 0
    cveNumber: str = ""
    severityId: int = 0
    severity: str = ""
    impactId: int = 0
    impact: str = ""
    issuingCna: str = ""
    initialReleaseDate: str = ""
    baseScore: str = ""
    temporalScore: str = ""
    vectorString: str = ""
    kbArticles: list[KbArticle]


class AffectedProductResponse(BaseModel):
    context: str = Field(alias='@odata.context', default="")
    count: str = Field(alias='@odata.count', default="")
    value: list[AffectedProduct]
    nextLink: str = Field(alias='@odata.nextLink', default="")
