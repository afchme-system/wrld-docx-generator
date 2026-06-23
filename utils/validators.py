import logging

logger = logging.getLogger(__name__)

def validate_request(template_name, content, output_filename):
    """
    Validate fill-template request parameters
    
    Args:
        template_name: Name of template file
        content: Dictionary with placeholder values
        output_filename: Output filename
    
    Raises:
        ValueError if validation fails
    """
    
    # Check template_name
    if not template_name:
        raise ValueError("template_name is required")
    
    if not isinstance(template_name, str):
        raise ValueError("template_name must be a string")
    
    if len(template_name) > 255:
        raise ValueError("template_name must be less than 255 characters")
    
    # Check content
    if not isinstance(content, dict):
        raise ValueError("content must be a dictionary")
    
    if len(content) == 0:
        logger.warning("content is empty - template will be returned unfilled")
    
    # Validate content values
    for key, value in content.items():
        if not isinstance(key, str):
            raise ValueError(f"content keys must be strings, got {type(key).__name__}")
        
        # Allow None, strings, numbers, and booleans
        if value is not None and not isinstance(value, (str, int, float, bool)):
            raise ValueError(
                f"content['{key}'] value must be string, number, boolean, or null, "
                f"got {type(value).__name__}"
            )
        
        # Warn if string is very large
        if isinstance(value, str) and len(value) > 5000:
            logger.warning(f"content['{key}'] is very large ({len(value)} chars)")
    
    # Check output_filename
    if not output_filename:
        raise ValueError("output_filename is required")
    
    if not isinstance(output_filename, str):
        raise ValueError("output_filename must be a string")
    
    if not output_filename.lower().endswith('.docx'):
        logger.warning("output_filename should end with .docx")
    
    if len(output_filename) > 255:
        raise ValueError("output_filename must be less than 255 characters")
    
    logger.info(
        f"Request validated: template={template_name}, "
        f"content_keys={len(content)}, output={output_filename}"
    )
