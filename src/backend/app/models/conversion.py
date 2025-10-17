# Conversion Job MongoDB Model
# Manages conversion process tracking and results
# - Job ID and user association
# - Source Excel file and target PowerPoint template
# - Conversion parameters and options
# - Processing status (pending, processing, completed, failed)
# - Progress tracking and estimated completion time
# - Error logs and debugging information
# - Output file details and download links
# PowerPoint Template and Conversion Configuration Models
# Handles template management, conversion settings, and processing analytics
# - Template definitions and access control
# - Conversion configuration and layout options
# - Processing analytics and performance tracking
# - Excel-to-PPT conversion workflow support


from datetime import datetime,timedelta
from typing import Optional,List,Dict,Any,Union
from enum import Enum
from pydantic import BaseModel,Field
from bson import ObjectId

# Import PyObjectId and TemplateCategory from user model
from .user import PyObjectId, TemplateCategory

class LayoutType(str,Enum):
    """Slide layout types"""
    TITLE_SLIDE="title_slide"
    TITLE_CONTENT="title_content"
    SECTION_HEADER="section_header"
    TWO_CONTENT="two_content"
    COMPARISON="comparison"
    TITLE_ONLY="title_only"
    BLANK="blank"
    CONTENT_WITH_CAPTION="content_with_caption"
    PICTURE_WITH_CAPTION="picture_with_caption"


class CharType(str,Enum):
    """Supoorted chart types for conversion"""
    BAR="bar"
    LINE="line"
    PIE="pie"
    COLUMN="column"
    AREA="area"
    SCATTER="scatter"
    DONUT="donut"
    COMBO="combo"


class DataMappingType(str,Enum):
    """How Excel data should mapped slides"""
    ROW_TO_SLIDE="row_to_slide"
    SHEET_TO_SLIDE="sheet_to_slide"
    CHART_TO_SLIDE="chart_to_slide"
    TABLE_TO_SLIDE="table_to_slide"
    SUMMARY_TO_SLIDE="summary_to_slide"

class PowerPointTemplate(BaseModel):
    """POwer template Defenations"""
    id:Optional[str]=Field(default=None,alias="_id")
    template_name:str=Field(...,min_length=1,max_length=100)
    category:TemplateCategory
    description:Optional[str]=Field(default="",max_length=500)

    # Template file info
    template_file_path:str
    thumbnail_path:Optional[str]=None
    preview_images:List[str]=Field(default_factory=list)

    # Template metdata
    slide_count:int=Field(ge=1,le=50)
    supported_layouts:List[LayoutType]=Field(default_factory=list)
    color_scheme:List[str]=Field(default_factory=list)
    font_families:List[str]=Field(default_factory=list)

    # Access control
    is_premium:bool=Field(default=False)
    required_subscription:List[str]=Field(default_factory=lambda:["free","pro","enterprise"])
    # Template configuartion
    default_settings:Dict[str,Any]=Field(default_factory=dict)
    customizable_elements:List[str]=Field(default_factory=list)


    #usage tracking
    is_active:bool=Field(default=True)
    created_by:Optional[PyObjectId]=None
    created_at:datetime=Field(default_factory=datetime.utcnow)
    updated_at:datetime=Field(default_factory=datetime.utcnow)

    class Config:
        populate_by_name=True
        json_encoders={
            ObjectId:str,
            datetime:lambda v:v.isoformat()
        }

class ConversionSettings(BaseModel):
    """user-specific conversion configurations"""
    #Layout preferences
    preferred_layouts:LayoutType=LayoutType.TITLE_CONTENT
    slide_orientation:str=Field(default="landscape",pattern="^(landscape|portrait)$")
    
    #data mapping configurations
    data_mapping_strategy:DataMappingType=DataMappingType.ROW_TO_SLIDE
    max_rows_per_slide:int=Field(default=10,ge=1,le=50)
    include_header:bool=Field(default=True)
    auto_format_numbers:bool=Field(default=True)

    # chart conversion settings
    convert_charts:bool=Field(default=True)
    chart_style:str=Field(default="modern")
    chart_color_scheme:str=Field(default="corporate")
    supported_chart_types:List[CharType]=Field(default_factory=lambda:[CharType.BAR,CharType.LINE,CharType.PIE])

    # Text Formatting
    font_size_title:int=Field(default=24,ge=12,le=48)
    font_size_body:int=Field(default=18,ge=8,le=36)
    font_family:str=Field(default="Calibri")

    # color and styling
    primary_color: str = Field(default="#1f4e79", pattern="^#[0-9a-fA-F]{6}$")
    secondary_color: str = Field(default="#70ad47", pattern="^#[0-9a-fA-F]{6}$")
    background_color: str = Field(default="#ffffff", pattern="^#[0-9a-fA-F]{6}$")


    # Advance options
    include_animations: bool = Field(default=False)
    compress_images: bool = Field(default=True)
    quality_level: str = Field(default="high", pattern="^(low|medium|high|ultra)$")

    class Config:
        json_schema_extra = {
            "example": {
                "preferred_layout": "title_content",
                "data_mapping_strategy": "row_to_slide",
                "max_rows_per_slide": 8,
                "convert_charts": True,
                "font_size_title": 28,
                "font_size_body": 20,
                "primary_color": "#1f4e79"
            }
        }


class ConversionResult(BaseModel):
    """Result of a conversion process"""
    conversion_id:str
    job_id:str
    user_id:PyObjectId

    #INput info
    source_file_id:str
    template_id:Optional[str]=None

    # Output info
    output_file_path:str
    output_file_url:Optional[str]=None
    output_file_size:int=Field(ge=0)
    slide_count:int=Field(ge=0)

    #Processing metrics
    processing_time_seconds:float=Field(ge=0.0)
    data_rows_processed:int=Field(default=0,ge=0)
    charts_converted:int=Field(default=0,ge=0)
    images_processed:int=Field(default=0,ge=0)

    # Quailty metrics
    success_rate:float=Field(default=100.0,ge=0.0,le=100.0)
    error_count:int=Field(default=0,ge=0)
    warning_count:int=Field(default=0,ge=0)

    # User Feedback
    user_rating:Optional[int]=Field(default=None,ge=1,le=5)
    user_feedback:Optional[str]=Field(default=None,max_length=1000)

    # Time stamp
    started_at:datetime
    completed_at:datetime
    expires_at:datetime=Field(default_factory=lambda:datetime.utcnow()+timedelta(hours=24))

    class Config:
        populate_by_name=True
        json_encoders={
            ObjectId:str,
            datetime:lambda v:v.isoformat()
        }



class TemplateUsageStats(BaseModel):
    """Tracks usage statistics for PowerPoint templates"""
    template_id:str
    template_name:str
    category:TemplateCategory

    # Usage metrics
    total_uses:int=Field(default=0,ge=0)
    unique_users:int=Field(default=0,ge=0)
    success_rate:float=Field(default=0,ge=0.0,le=100.0)

    # User feedback
    average_rating:Optional[float]=Field(default=None,ge=1.0,le=5.0)
    total_ratings:int=Field(default=0,ge=0)
    

    # Trends (Last 30 days)
    usage_trend:float=Field(default=0.0)
    rating_trend:float=Field(default=0.0)

# Utility functions for conversion logic
def get_template_by_category(category:TemplateCategory,user_subscription:str)->List[str]:
    """Get available templates for user based on subscriptions"""
    pass


def calculate_conversion_credits(file_size_mb:float,slide_count:int,template_category:TemplateCategory,include_charts:bool=False)->int:
    """Calculate credits needed for conversion based on parameters"""
    base_credits=1
    
    #File size factor(>10MB +1 credit, >50MB +2 credits)
    if file_size_mb>10:
        base_credits+=1
    if file_size_mb>50:
        base_credits+=1

    # Slide count factor(>10 slides +1 credit)
    if slide_count>10:
        base_credits+=1

    # Premium template factor
    if template_category in [TemplateCategory.PREMIUM,TemplateCategory.CUSTOM]:
        base_credits+=2
    elif template_category==TemplateCategory.PROFESSIONAL:
        base_credits+=1

    # Chart Conversion Factor
    if include_charts:
        base_credits+=1

    return min(base_credits,10)
    
def validate_conversion_settings(settings:ConversionSettings)->Dict[str,Any]:
    """Validate user conversion settings and return any issues"""
    validate_results={"is_valid":True,"errors":[],"warnings":[],"normalized_settings":settings.model_dump()}

    # add validation logic here
    # Check colour format ,font availability,etc.
    return validate_results

def estimate_conversion_time(file_size_mb:float,row_count:int,chart_count:int,template_complexity:str="medium")->float:
    """Estimate conversion time in seconds based on input parameters"""
    base_time=100
    # file size factor
    base_time+=file_size_mb*2

    # Row processing
    base_time+=row_count*0.1

    # Chart processing
    base_time+=chart_count*5

    # Template complexity factor
    complexity_multipliers={
        "simple":0.8,
        "medium":1.0,
        "complex":1.2,
        "premium":2.0,
    }.get(template_complexity,1.0)

    return base_time*complexity_multipliers