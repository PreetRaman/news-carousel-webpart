# News Carousel Web Part

## Summary

The News Carousel Web Part is a SharePoint Framework (SPFx) solution that displays news articles from your SharePoint site in an interactive carousel format. It automatically fetches the latest news articles from the "Site Pages" list and presents them in a visually appealing, rotating carousel with navigation controls and auto-play functionality.

**Technologies Used:**
- SharePoint Framework 1.21.1
- React 17.0.1
- TypeScript 5.3.3
- Fluent UI (Office UI Fabric)
- SharePoint REST API

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- SharePoint Online
- Microsoft Teams

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js >= 18.17.0 < 19.0.0
- SharePoint Online site with news articles (Site Pages list with ContentType "Article")
- SPFx development environment configured
- Microsoft 365 developer tenant (for testing)

## Solution

| Solution      | Author(s)                      |
| ------------- | ------------------------------ |
| news-carousel | Ramanpreet Kaur             |

## Version history

| Version | Date             | Comments                    |
| ------- | ---------------- | --------------------------- |
| 1.0     | October 15, 2025 | Initial release             |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - **npm install**
  - **gulp serve**
- Open the SharePoint workbench and add the News Carousel web part

## Testing the Web Part

### Prerequisites for Testing

Before testing, ensure you have:
1. **SharePoint Site with News Articles**: Your SharePoint site must have at least 3-5 news articles published
   - Go to your SharePoint site
   - Create news articles (Site Pages with ContentType "Article")
   - Ensure articles have titles, descriptions, and are published

2. **Update serve.json** (if needed): 
   - Open `config/serve.json`
   - Update the `initialPage` URL to point to your SharePoint site:
   ```json
   {
     "initialPage": "https://yourtenant.sharepoint.com/sites/yoursite"
   }
   ```

### Method 1: Local Workbench Testing (Recommended for Development)

1. **Start the local development server**:
   ```bash
   npm install
   gulp serve
   ```

2. **Open the local workbench**:
   - The command will automatically open `https://localhost:4321/temp/workbench.html`
   - Or manually navigate to: `https://localhost:4321/temp/workbench.html`

3. **Add the web part**:
   - Click the **+** button to add a web part
   - Select **News Carousel** from the web part picker

4. **Configure the web part**:
   - Click the edit (pencil) icon on the web part
   - Configure the properties:
     - **Title**: Enter a title (e.g., "Latest News")
     - **Number of items to show**: Set to 3-5
     - **Show navigation arrows**: Toggle ON/OFF
     - **Auto-play carousel**: Toggle ON/OFF
     - **Auto-play interval**: Set milliseconds (e.g., 5000 for 5 seconds)

5. **Note**: The local workbench may show limited data or mock data. For full functionality, use the hosted workbench.

### Method 2: Hosted Workbench Testing (Recommended for Full Testing)

1. **Start the development server**:
   ```bash
   gulp serve
   ```

2. **Open the hosted workbench**:
   - Navigate to: `https://yourtenant.sharepoint.com/sites/yoursite/_layouts/workbench.aspx`
   - Replace `yourtenant` and `yoursite` with your actual SharePoint site URL

3. **Add and test the web part**:
   - Click the **+** button
   - Select **News Carousel**
   - The web part will fetch real news articles from your SharePoint site

4. **Test all features**:
   - Verify news articles are displayed
   - Test navigation arrows (if enabled)
   - Click on news items to verify they open in a new tab
   - Test dot navigation indicators
   - Test auto-play functionality (if enabled)

### Method 3: Testing in a SharePoint Page

1. **Build the solution**:
   ```bash
   gulp build
   ```

2. **Bundle and package**:
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

3. **Deploy the package**:
   - Go to your SharePoint Admin Center
   - Navigate to **More features** → **Apps** → **App Catalog**
   - Upload the `.sppkg` file from the `sharepoint/solution` folder
   - Deploy the solution

4. **Add to a page**:
   - Go to any SharePoint page
   - Click **Edit**
   - Click **+** to add a web part
   - Select **News Carousel** from the web part picker
   - Configure and test

### What to Test

#### ✅ Functional Testing Checklist

- [ ] **Data Loading**: Verify news articles load from SharePoint
- [ ] **Display**: Confirm 3 items are visible (previous, current, next)
- [ ] **Navigation**: 
  - [ ] Left arrow navigates to previous item
  - [ ] Right arrow navigates to next item
  - [ ] Dot indicators allow direct navigation
  - [ ] Navigation loops from last to first item (circular)
- [ ] **Click Functionality**: Clicking a news item opens the article in a new tab
- [ ] **Auto-Play**: 
  - [ ] Carousel auto-advances when enabled
  - [ ] Auto-play stops when navigating manually (if implemented)
- [ ] **Configuration**: 
  - [ ] Title displays correctly
  - [ ] Arrows show/hide based on setting
  - [ ] Auto-play respects interval setting
- [ ] **Error Handling**: 
  - [ ] Loading state displays while fetching data
  - [ ] Error message shows if data fetch fails
  - [ ] Empty state shows when no news articles exist
- [ ] **Responsive Design**: 
  - [ ] Carousel works on different screen sizes
  - [ ] Layout adapts to mobile/tablet/desktop

#### 🔍 Browser Console Testing

1. **Open Developer Tools** (F12)
2. **Check for errors**: Look for any JavaScript errors in the console
3. **Network tab**: Verify REST API calls to SharePoint are successful
4. **Check API response**: Ensure news articles are being fetched correctly

### Troubleshooting

**Issue: No news articles are displayed**
- Verify your SharePoint site has news articles with ContentType "Article"
- Check browser console for API errors
- Ensure you're testing in the hosted workbench (not local workbench)
- Verify you have proper permissions to read the Site Pages list

**Issue: Web part shows "Loading..." indefinitely**
- Check browser console for errors
- Verify SharePoint REST API endpoint is accessible
- Check network tab for failed API requests
- Ensure `NewsService.ts` is correctly configured

**Issue: Navigation arrows don't work**
- Verify `showArrows` property is set to `true`
- Check browser console for JavaScript errors
- Ensure click handlers are properly bound

**Issue: Auto-play not working**
- Verify `autoPlay` property is set to `true`
- Check that `autoPlayInterval` is set to a valid number (milliseconds)
- Look for console errors related to `setInterval`

**Issue: Cannot add web part to page**
- Ensure the solution is properly built and packaged
- Verify the package is deployed to the App Catalog
- Check that the web part manifest is correctly configured

### Debugging Tips

1. **Use VS Code Debugger**: 
   - The `.vscode/launch.json` file is configured for debugging
   - Set breakpoints in your TypeScript files
   - Press F5 to start debugging

2. **Browser DevTools**:
   - Use React DevTools extension to inspect component state
   - Monitor network requests to SharePoint API
   - Check console for detailed error messages

3. **Check Service Implementation**:
   - Verify `NewsService.ts` is making correct API calls
   - Test the REST API endpoint directly in browser
   - Ensure data mapping is correct in `mapToNewsItems()`

## Features

### Core Functionality

- **Automatic News Fetching**: Retrieves the latest news articles from SharePoint Site Pages list
- **Interactive Carousel**: Displays 3 news items at a time (previous, current, next) with smooth transitions
- **Navigation Controls**:
  - Left/Right arrow buttons (optional)
  - Dot indicators for direct navigation
  - Click on news items to open the full article
- **Auto-Play**: Configurable automatic rotation of news items
- **Responsive Design**: Adapts to different screen sizes and themes
- **Theme Support**: Automatically adapts to SharePoint dark/light themes

### Configuration Options

The web part offers the following customizable properties:

- **Title**: Custom title for the carousel section
- **Number of Items to Show**: Slider control (3-5 items)
- **Show Navigation Arrows**: Toggle to show/hide arrow buttons
- **Auto-Play**: Enable/disable automatic rotation
- **Auto-Play Interval**: Time in milliseconds between slides (when auto-play is enabled)

### Data Displayed

Each news item in the carousel displays:
- Article thumbnail/image
- Title
- Description (truncated to 150 characters)
- Author name
- Published date
- Clickable link to the full article

## Architecture

### Components

1. **NewsCarouselWebPart.ts**: Main web part class that handles property pane configuration and rendering
2. **NewsCarousel.tsx**: Main React component that manages carousel state and navigation
3. **NewsItem.tsx**: Individual news item component that displays article information
4. **NewsService.ts**: Service class that fetches news articles from SharePoint REST API
5. **iNewsItem.ts**: TypeScript interface defining the news item data structure

### How It Works

1. **Data Fetching**: 
   - The `NewsService` connects to SharePoint REST API
   - Queries the "Site Pages" list for items with ContentType "Article"
   - Retrieves the 10 most recent articles sorted by Modified date
   - Maps SharePoint list items to `INewsItem` objects

2. **Display Logic**:
   - Shows 3 items simultaneously: previous, current (active), and next
   - The active item is highlighted; others are dimmed
   - Smooth circular navigation (loops from last to first item)

3. **Navigation**:
   - Users can navigate using arrow buttons or dot indicators
   - Clicking a news item opens the article in a new browser tab
   - Auto-play automatically advances slides at configured intervals

4. **State Management**:
   - Tracks current index, loading state, error state, and news items array
   - Handles cleanup of auto-play timers on component unmount

### SharePoint Integration

- Uses `SPHttpClient` to make authenticated REST API calls
- Filters news articles by ContentType "Article"
- Supports SharePoint theme variables for consistent styling
- Works seamlessly in SharePoint Online and Microsoft Teams

This extension illustrates the following concepts:

- SharePoint Framework web part development
- React component state management
- SharePoint REST API integration
- SharePoint property pane configuration
- Theme-aware styling with SCSS modules
- TypeScript interfaces and type safety
- Responsive UI design patterns

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
