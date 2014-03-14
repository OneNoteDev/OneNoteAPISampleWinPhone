//*********************************************************
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// Licensed under the Apache License, Version 2.0 (the ""License"");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// http://www.apache.org/licenses/LICENSE-2.0
//
// THIS CODE IS PROVIDED ON AN  *AS IS* BASIS, WITHOUT
// WARRANTIES OR CONDITIONS OF ANY KIND, EITHER EXPRESS
// OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
// WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR
// PURPOSE, MERCHANTABLITY OR NON-INFRINGEMENT.
//
// See the Apache Version 2.0 License for specific language
// governing permissions and limitations under the License.
//*********************************************************

using Microsoft.Live;
using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Resources;

/********************************************************
 * 
 * README
 * 
 * 1. Before hitting F5, please ensure latest nuget update is installed and package restore is enabled in tools/options/package manager
 * 2. Before re-using the sample code in your app, please replace sample ClientId on MainPage.xaml with your own ClientId
 * 3. Accessing the OneNote Service APIs only requires the Office.onenote_* scopes. This sample app also contains the wl.offline_access wl.signin
 *   scopes in order to demonstrate how to use a refresh token and re-authorize access tokens.
 * *****************************************************/
namespace WinPhoneOneNoteServiceSample
{
    public partial class MainPage
    {
        private string _accessToken;
        private DateTimeOffset _accessTokenExpiration;
        private string _refreshToken; // Refresh token (only applicable when the app uses the wl.offline_access wl.signin scopes)
        private StandardResponse _response;

        // OneNote Service API v1.0 Endpoint
        private const string PagesEndpoint = "https://www.onenote.com/api/v1.0/pages";

        // Collateral used to refresh access token (only applicable when the app uses the wl.offline_access wl.signin scopes)
        private const string MsaTokenRefreshUrl = "https://login.live.com/oauth20_token.srf";
        private const string TokenRefreshContentType = "application/x-www-form-urlencoded";
        private const string TokenRefreshRedirectUri = "https://login.live.com/oauth20_desktop.srf";
        private const string TokenRefreshRequestBody = "client_id={0}&redirect_uri={1}&grant_type=refresh_token&refresh_token={2}";

        // Constructor
        public MainPage()
        {
            InitializeComponent();
            this.Loaded += CheckIfClientIdUpdated;
        }

        private void CheckIfClientIdUpdated(object sender, RoutedEventArgs e)
        {
            if ((string)Resources["ClientId"] == "Insert Your Client Id Here")
            {
                MessageBox.Show("Visit http://go.microsoft.com/fwlink/?LinkId=392537 for instructions on getting a Client Id. Please specify your client ID in the MainPage.xaml file and rebuild the application.");
            }
        }

        #region Send create page requests

        /// <summary>
        /// Send a create page request
        /// </summary>
        /// <param name="createMessage">The HttpRequestMessage which contains the page information</param>
        private async Task SendCreatePageRequest(HttpRequestMessage createMessage)
        {
            var httpClient = new HttpClient();
            _response = null;

            // Check if Auth token needs to be refreshed
            await RefreshAuthTokenIfNeeded();

            // Add Authorization header
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);
            // Note: API only supports JSON return type.
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            // Get and parse the HTTP response from the service
            HttpResponseMessage response = await httpClient.SendAsync(createMessage);
            _response = await ParseResponse(response);

            // Update the UI accordingly
            UpdateUIAfterRequest(response);
        }

        /// <summary>
        /// Create a page with simple html content
        /// </summary>
        private async void btn_CreateSimple_Click(object sender, RoutedEventArgs e)
        {
            string date = GetDate();

            string simpleHtml = "<html>" +
                          "<head>" +
                          "<title>A page created from basic HTML-formatted text (WinPhone8 Sample)</title>" +
                          "<meta name=\"created\" content=\"" + date + "\" />" +
                          "</head>" +
                          "<body>" +
                          "<p>This is a page that just contains some simple <i>formatted</i> <b>text</b></p>" +
                          "</body>" +
                          "</html>";

            // Create the request message, which is a text/html single part in this case
            // The Service also supports content type multipart/form-data for more complex scenarios
            var createMessage = new HttpRequestMessage(HttpMethod.Post, PagesEndpoint)
            {
                Content = new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html")
            };

            await SendCreatePageRequest(createMessage);
        }

        /// <summary>
        /// Create page with image
        /// </summary>
        private async void btn_CreateWithImage_Click(object sender, RoutedEventArgs e)
        {
            const string imagePartName = "image1";

            string date = GetDate();
            string simpleHtml = "<html>" +
                                "<head>" +
                                "<title>A page with an image on it (WinPhone8 Sample)</title>" +
                                "<meta name=\"created\" content=\"" + date + "\" />" +
                                "</head>" +
                                "<body>" +
                                "<h1>This is a page with an image on it</h1>" +
                                "<img src=\"name:" + imagePartName + "\" alt=\"A beautiful logo\" width=\"426\" height=\"68\" />" +
                                "</body>" +
                                "</html>";

            // Create the image part - make sure it is disposed after we've sent the message in order to close the stream.
            Stream imageStream = GetAssetFileStream("assets\\Logo.jpg");
            using (var imageContent = new StreamContent(imageStream))
            {
                imageContent.Headers.ContentType = new MediaTypeHeaderValue("image/jpeg");
                HttpRequestMessage createMessage = new HttpRequestMessage(HttpMethod.Post, PagesEndpoint)
                {
                    // Create a multipart/form data request in this case
                    // The Service also supports single part text/html content for simple scenarios
                    Content = new MultipartFormDataContent
                    {
                        {new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html"), "Presentation"},
                        {imageContent, imagePartName}
                    }
                };

                // Must send the request within the using block, or the image stream will have been disposed.
                await SendCreatePageRequest(createMessage);
            }
        }

        /// <summary>
        /// Create page with URL snapshot
        /// </summary>
        private async void btn_CreateWithUrl_Click(object sender, RoutedEventArgs e)
        {
            string date = GetDate();
            string simpleHtml = @"<html>" +
                                "<head>" +
                                "<title>A Page Created With a URL Snapshot on it (WinPhone8 Sample)</title>" +
                                "<meta name=\"created\" content=\"" + date + "\" />" +
                                "</head>" +
                                "<body>" +
                                "<p>This is a page with an image of an html page rendered from a URL on it.</p>" +
                                "<img data-render-src=\"http://www.onenote.com\" alt=\"An important web page\"/>" +
                                "</body>" +
                                "</html>";

            // Create the request message, which is a text/html single part in this case
            // The Service also supports content type multipart/form-data for more complex scenarios
            var createMessage = new HttpRequestMessage(HttpMethod.Post, PagesEndpoint)
            {
                Content = new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html")
            };

            await SendCreatePageRequest(createMessage);
        }

        /// <summary>
        /// Create page with HTML snapshot
        /// </summary>
        private async void btn_CreateWithHtml_Click(object sender, RoutedEventArgs e)
        {
            const string embeddedPartName = "embedded1";
            const string embeddedWebPage =
                "<html>" +
                "<head>" +
                "<title>Embedded HTML</title>" +
                "</head>" +
                "<body>" +
                "<h1>This is a screen grab of a web page</h1>" +
                "<p>Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nullam vehicula magna quis mauris accumsan, nec imperdiet nisi tempus. Suspendisse potenti. " +
                "Duis vel nulla sit amet turpis venenatis elementum. Cras laoreet quis nisi et sagittis. Donec euismod at tortor ut porta. Duis libero urna, viverra id " +
                "aliquam in, ornare sed orci. Pellentesque condimentum gravida felis, sed pulvinar erat suscipit sit amet. Nulla id felis quis sem blandit dapibus. Ut " +
                "viverra auctor nisi ac egestas. Quisque ac neque nec velit fringilla sagittis porttitor sit amet quam.</p>" +
                "</body>" +
                "</html>";

            string date = GetDate();

            string simpleHtml = "<html>" +
                                "<head>" +
                                "<title>A Page Created With Snapshot of Webpage in it (WinPhone8 Sample)</title>" +
                                "<meta name=\"created\" content=\"" + date + "\" />" +
                                "</head>" +
                                "<body>" +
                                "<h1>This is a page with an image of an html page on it.</h1>" +
                                "<img data-render-src=\"name:" + embeddedPartName + "\" alt=\"A website screen grab\" />" +
                                "</body>" +
                                "</html>";

            var createMessage = new HttpRequestMessage(HttpMethod.Post, PagesEndpoint)
            {
                // Create a multipart/form data request in this case
                // The Service also supports single part text/html for more simpler scenarios
                Content = new MultipartFormDataContent
                        {
                            {new StringContent(simpleHtml, System.Text.Encoding.UTF8, "text/html"), "Presentation"},
                            {new StringContent(embeddedWebPage, System.Text.Encoding.UTF8, "text/html"), embeddedPartName}
                        }
            };

            await SendCreatePageRequest(createMessage);
        }

        /// <summary>
        /// Creates a OneNote page with a file attachment
        /// </summary>
        private async void btn_CreateWithAttachment_Click(object sender, RoutedEventArgs e)
        {
            const string attachmentPartName = "pdfattachment1";
            string date = GetDate();
            string attachmentHtml = "<html>" +
                                "<head>" +
                                "<title>A page created with a file attachment (WinPhone8 Sample)</title>" +
                                "<meta name=\"created\" content=\"" + date + "\" />" +
                                "</head>" +
                                "<body>" +
                                "<h1>This is a page with a pdf file attachment</h1>" +
                                "<object data-attachment=\"attachment.pdf\" data=\"name:" + attachmentPartName + "\" />" +
                                "</body>" +
                                "</html>";
            // Create the attachment part - make sure it is disposed after we've sent the message in order to close the stream.
            Stream attachmentStream = GetAssetFileStream("Assets\\attachment.pdf");
            using (var attachmentContent = new StreamContent(attachmentStream))
            {
                attachmentContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                HttpRequestMessage createMessage = new HttpRequestMessage(HttpMethod.Post, PagesEndpoint)
                {
                    Content = new MultipartFormDataContent
                            {
                                {new StringContent(attachmentHtml, System.Text.Encoding.UTF8, "text/html"), "Presentation"},
                                {attachmentContent, attachmentPartName}
                            }
                };
                // Must send the request within the using block, or the attachment file stream will have been disposed.
                await SendCreatePageRequest(createMessage);
            }
        }

        #endregion

        #region Open OneNote Hyperlink

        /// <summary>
        /// Open the created page in the OneNote app
        /// </summary>
        private async void HyperlinkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_response as CreateSuccessResponse != null)
            {
                CreateSuccessResponse successResponse = (CreateSuccessResponse)_response;
                await Windows.System.Launcher.LaunchUriAsync(FormulatePageUri(successResponse.OneNoteClientUrl));
            }
        }

        /// <summary>
        /// Formulate the OneNoteClientUrl so that we can open the OneNote app directly
        /// </summary>
        /// <param name="oneNoteClientUrl">The OneNoteClientUrl received in the API JSON response</param>
        private static Uri FormulatePageUri(string oneNoteClientUrl)
        {
            // Regular expression for identifying GUIDs in the URL returned by the server.
            // We need to wrap such GUIDs in curly braces before sending them to OneNote.
            Regex guidRegex = new Regex(@"=([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})&",
                RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
            if (!String.IsNullOrEmpty(oneNoteClientUrl))
            {
                var matches = guidRegex.Matches(oneNoteClientUrl);
                if (matches.Count == 2)
                {
                    oneNoteClientUrl =
                        oneNoteClientUrl.Replace(matches[0].Groups[1].Value, "{" + matches[0].Groups[1].Value + "}")
                            .Replace(matches[1].Groups[1].Value, "{" + matches[1].Groups[1].Value + "}");
                }
                return new Uri(oneNoteClientUrl);
            }
            return null;
        }

        #endregion

        #region LiveSDK authentication/token refresh

        /// <summary>
        /// This method is called when the Live session status changes
        /// </summary>
        private void OnSessionChanged(object sender, Microsoft.Live.Controls.LiveConnectSessionChangedEventArgs e)
        {
            switch (e.Status)
            {
                case LiveConnectSessionStatus.Connected:
                    _accessToken = e.Session.AccessToken;
                    _accessTokenExpiration = e.Session.Expires;
                    _refreshToken = e.Session.RefreshToken;

                    SetCreateButtonsEnabled(true);
                    infoTextBlock.Text = "Authentication successful";
                    break;
                case LiveConnectSessionStatus.NotConnected:
                    SetCreateButtonsEnabled(false);
                    infoTextBlock.Text = "Authentication failed.";
                    break;
                default:
                    SetCreateButtonsEnabled(false);
                    infoTextBlock.Text = "Not Authenticated";
                    break;
            }
        }

        /// <summary>
        /// This method tries to refresh the token if it expires. The authentication token needs to be
        /// refreshed continuosly, so that the user is not prompted to sign in again
        /// </summary>
        /// <returns></returns>
        private async Task AttemptAccessTokenRefresh()
        {
            var createMessage = new HttpRequestMessage(HttpMethod.Post, MsaTokenRefreshUrl)
            {
                Content = new StringContent(
                    String.Format(CultureInfo.InvariantCulture, TokenRefreshRequestBody,
                        Resources["ClientId"],
                        TokenRefreshRedirectUri,
                        _refreshToken),
                    System.Text.Encoding.UTF8,
                    TokenRefreshContentType)
            };

            HttpClient httpClient = new HttpClient();
            HttpResponseMessage response = await httpClient.SendAsync(createMessage);
            await ParseRefreshTokenResponse(response);
        }

        /// <summary>
        /// Handle the RegreshToken response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the TokenRefresh request</param>
        private async Task ParseRefreshTokenResponse(HttpResponseMessage response)
        {
            if (response.StatusCode == HttpStatusCode.OK)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                _accessToken = responseObject.access_token;
                _accessTokenExpiration = _accessTokenExpiration.AddSeconds((double)responseObject.expires_in);
                _refreshToken = responseObject.refresh_token;
            }
        }

        #endregion

        #region Helper functions

        /// <summary>
        /// Get date in ISO8601 format with local timezone offset
        /// </summary>
        /// <returns>Date as ISO8601 string</returns>
        private static string GetDate()
        {
            return DateTime.Now.ToString("o");
        }

        /// <summary>
        /// Get an asset file packaged with the application and return it as a managed stream
        /// </summary>
        /// <param name="assetFile">The path name of an asset relative to the application package root</param>
        /// <returns>A managed stream of the asset file data, opened for read</returns>
        private static Stream GetAssetFileStream(string assetFile)
        {
            StreamResourceInfo resource = Application.GetResourceStream(new Uri(assetFile, UriKind.Relative));
            return resource.Stream;
        }

        /// <summary>
        /// Refreshes the live authentication token if it has expired
        /// </summary>
        private async Task RefreshAuthTokenIfNeeded()
        {
            if (_accessTokenExpiration.CompareTo(DateTimeOffset.UtcNow) <= 0)
            {
                infoTextBlock.Text = "Access token needs to be refreshed";
                await AttemptAccessTokenRefresh();
            }
            infoTextBlock.Text = "Sending Request...";
            hyperlinkCreatedPage.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Update the UI after a create page request, depending on if it was
        /// successful or not
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private void UpdateUIAfterRequest(HttpResponseMessage response)
        {
            if (response.StatusCode == HttpStatusCode.Created)
            {
                infoTextBlock.Text = "Page successfully created.";
                hyperlinkCreatedPage.Visibility = Visibility.Visible;
            }
            else
            {
                infoTextBlock.Text = "Page creation failed with error code: " + response.StatusCode;
                hyperlinkCreatedPage.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// Parse the OneNote Service API create page response
        /// </summary>
        /// <param name="response">The HttpResponseMessage from the create page request</param>
        private async static Task<StandardResponse> ParseResponse(HttpResponseMessage response)
        {
            StandardResponse standardResponse;
            if (response.StatusCode == HttpStatusCode.Created)
            {
                dynamic responseObject = JsonConvert.DeserializeObject(await response.Content.ReadAsStringAsync());
                standardResponse = new CreateSuccessResponse
                {
                    StatusCode = response.StatusCode,
                    OneNoteClientUrl = responseObject.links.oneNoteClientUrl.href,
                    OneNoteWebUrl = responseObject.links.oneNoteWebUrl.href
                };
            }
            else
            {
                standardResponse = new StandardErrorResponse
                {
                    StatusCode = response.StatusCode,
                    Message = await response.Content.ReadAsStringAsync()
                };
            }

            // Extract the correlation id.  Apps should log this if they want to collect data to diagnose failures with Microsoft support 
            IEnumerable<string> correlationValues;
            if (response.Headers.TryGetValues("X-CorrelationId", out correlationValues))
            {
                standardResponse.CorrelationId = correlationValues.FirstOrDefault();
            }

            return standardResponse;
        }

        /// <summary>
        /// Change "create page" buttons to enabled/disabled
        /// </summary>
        /// <param name="shouldBeEnabled">A bool value indicating if the buttons should be enabled</param>
        private void SetCreateButtonsEnabled(bool shouldBeEnabled)
        {
            btn_CreateSimple.IsEnabled = shouldBeEnabled;
            btn_CreateWithHtml.IsEnabled = shouldBeEnabled;
            btn_CreateWithImage.IsEnabled = shouldBeEnabled;
            btn_CreateWithUrl.IsEnabled = shouldBeEnabled;
            btn_CreateWithAttachment.IsEnabled = shouldBeEnabled;
        }

        #endregion
    }
}