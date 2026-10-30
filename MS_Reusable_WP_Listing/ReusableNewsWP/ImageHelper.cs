using System;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Microsoft.SharePoint;

namespace MS_Reusable_WP_Listing.ReusableNewsWP
{
    /// <summary>
    /// Helper class for handling SharePoint image fields
    /// </summary>
    public static class ImageHelper
    {
        /// <summary>
        /// Extracts image URL from a SharePoint list item
        /// </summary>
        /// <param name="item">SharePoint list item</param>
        /// <param name="fieldName">Field name containing the image</param>
        /// <returns>URL of the image or placeholder if not found</returns>
        public static string GetImageUrl(SPListItem item, string fieldName)
        {
            if (item == null || string.IsNullOrEmpty(fieldName))
                return "/_layouts/15/images/placeholder.png";

            try
            {
                if (item.Fields.ContainsField(fieldName) && item[fieldName] != null)
                {
                    object valueObj = item[fieldName];
                    if (valueObj == null)
                        return "/_layouts/15/images/placeholder.png";

                    string fieldValueStr = valueObj.ToString();
                    SPField fieldObj = item.Fields.GetFieldByInternalName(fieldName);
                    string fieldTypeStr = fieldObj.GetType().Name;

                    // Debug information
                    Debug.WriteLine($"Field type: {fieldTypeStr}, Value: {fieldValueStr}");

                    // Handle special image fields
                    if (fieldTypeStr.Contains("Publishing") || fieldTypeStr.Contains("Image"))
                    {
                        string extractedUrl = ExtractSpecialImageUrl(fieldValueStr);
                        if (!string.IsNullOrEmpty(extractedUrl))
                            return extractedUrl;
                    }

                    // Handle URL field
                    if (fieldObj is SPFieldUrl)
                    {
                        try
                        {
                            SPFieldUrlValue urlValueObj = new SPFieldUrlValue(fieldValueStr);
                            return urlValueObj.Url;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error parsing URL field: {ex.Message}");
                        }
                    }

                    // Handle HTML content with image
                    if (fieldValueStr.Contains("src="))
                    {
                        Match srcMatchObj = Regex.Match(fieldValueStr, @"src=""([^""]+)""");
                        if (srcMatchObj.Success && srcMatchObj.Groups.Count > 1)
                        {
                            return srcMatchObj.Groups[1].Value;
                        }
                    }

                    // If it's a direct URL
                    if (fieldValueStr.StartsWith("http") || fieldValueStr.StartsWith("/"))
                    {
                        return fieldValueStr;
                    }

                    // Last resort: try to find any URL in the string
                    Match anyUrlMatchObj = Regex.Match(fieldValueStr, @"(https?://[^\s"",<>]+|/[^\s"",<>]+\.(jpg|jpeg|png|gif))");
                    if (anyUrlMatchObj.Success)
                    {
                        return anyUrlMatchObj.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in GetImageUrl: {ex.Message}");
            }

            return "/_layouts/15/images/placeholder.png";
        }

        /// <summary>
        /// Extracts URL from SharePoint special image field value
        /// </summary>
        /// <param name="fieldValueStr">The field value as string</param>
        /// <returns>URL of the image or null if not found</returns>
        public static string ExtractSpecialImageUrl(string fieldValueStr)
        {
            try
            {
                if (string.IsNullOrEmpty(fieldValueStr))
                    return null;

                // Try to extract the URL from the BlobUrl property
                if (fieldValueStr.Contains("BlobUrl"))
                {
                    Match blobMatchObj = Regex.Match(fieldValueStr, @"""BlobUrl"":""([^""]+)""");
                    if (blobMatchObj.Success && blobMatchObj.Groups.Count > 1)
                    {
                        string extractedUrl = blobMatchObj.Groups[1].Value;
                        // Fix escaped backslashes
                        extractedUrl = extractedUrl.Replace("\\\\", "\\");
                        Debug.WriteLine($"Extracted BlobUrl: {extractedUrl}");
                        return extractedUrl;
                    }
                }
                
                // Try to extract from special JSON format
                if (fieldValueStr.Contains("PublishingWebDatabase"))
                {
                    // This is likely a special image field
                    Debug.WriteLine("Detected special image field");
                    
                    // Try to find the image URL in the JSON
                    Match defaultUrlMatchObj = Regex.Match(fieldValueStr, @"""DefaultUrl"":""([^""]+)""");
                    if (defaultUrlMatchObj.Success && defaultUrlMatchObj.Groups.Count > 1)
                    {   
                        string extractedUrl = defaultUrlMatchObj.Groups[1].Value.Replace("\\\\", "\\");
                        Debug.WriteLine($"Extracted DefaultUrl: {extractedUrl}");
                        return extractedUrl;
                    }
                }

                // Try to extract from ImageUrl property
                if (fieldValueStr.Contains("ImageUrl"))
                {
                    Match imageUrlMatchObj = Regex.Match(fieldValueStr, @"""ImageUrl"":""([^""]+)""");
                    if (imageUrlMatchObj.Success && imageUrlMatchObj.Groups.Count > 1)
                    {
                        string extractedUrl = imageUrlMatchObj.Groups[1].Value;
                        // Fix escaped backslashes
                        extractedUrl = extractedUrl.Replace("\\\\", "\\");
                        return extractedUrl;
                    }
                }

                // Try to extract from server URL format
                if (fieldValueStr.Contains("\"serverUrl\"") && fieldValueStr.Contains("\"serverRelativeUrl\""))
                {
                    Match serverMatchObj = Regex.Match(fieldValueStr, @"""serverUrl"":""([^""]+)""");
                    Match relativeMatchObj = Regex.Match(fieldValueStr, @"""serverRelativeUrl"":""([^""]+)""");
                    
                    if (serverMatchObj.Success && relativeMatchObj.Success && 
                        serverMatchObj.Groups.Count > 1 && relativeMatchObj.Groups.Count > 1)
                    {
                        string serverUrlStr = serverMatchObj.Groups[1].Value.Replace("\\\\", "\\");
                        string relativeUrlStr = relativeMatchObj.Groups[1].Value.Replace("\\\\", "\\");
                        
                        // Combine URLs properly
                        if (relativeUrlStr.StartsWith("/"))
                        {
                            return serverUrlStr + relativeUrlStr;
                        }
                        else
                        {
                            return serverUrlStr + "/" + relativeUrlStr;
                        }
                    }
                }

                // Try to find any URL in the value
                Match generalUrlMatchObj = Regex.Match(fieldValueStr, @"https?://[^\s"",<>]+");
                if (generalUrlMatchObj.Success)
                {
                    return generalUrlMatchObj.Value;
                }

                // Try to find any relative URL for images
                Match relativeUrlMatchObj = Regex.Match(fieldValueStr, @"/[^\s"",<>]+\.(jpg|jpeg|png|gif|bmp)");
                if (relativeUrlMatchObj.Success)
                {
                    return relativeUrlMatchObj.Value;
                }
                
                // Special case for image field with specific format
                if (fieldValueStr.Contains("<img") && fieldValueStr.Contains("PublishingImages"))
                {
                    Match imgTagMatchObj = Regex.Match(fieldValueStr, @"<img[^>]*src=""([^""]*PublishingImages[^""]*)""");
                    if (imgTagMatchObj.Success && imgTagMatchObj.Groups.Count > 1)
                    {
                        string extractedImgUrl = imgTagMatchObj.Groups[1].Value;
                        Debug.WriteLine($"Extracted img tag URL: {extractedImgUrl}");
                        return extractedImgUrl;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error extracting image URL: {ex.Message}");
            }

            return null;
        }
    }
}
