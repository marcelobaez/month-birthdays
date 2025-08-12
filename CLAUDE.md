# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a SharePoint Framework (SPFx) 1.21.0 webpart that displays employee birthdays from a SharePoint list. The project is called "test-webpart-node-22" but contains a Month Birthdays webpart that reads birthday data from a SharePoint list called 'birthdays_eby' and displays upcoming birthdays for the current month.

## Development Commands
- `npm run serve` - Start development server using fast-serve for improved development experience
- `npm run build` - Build the webpart using gulp bundle
- `npm run clean` - Clean build artifacts using gulp clean  
- `npm run test` - Run tests using gulp test

Note: The project uses `fast-serve` instead of the traditional `gulp serve` for better development experience.

## Architecture

### Core Structure
- **Main webpart**: `src/webparts/monthBirthdaysWebpart/MonthBirthdaysWebpart.ts` - SPFx webpart entry point
- **React component**: `src/webparts/monthBirthdaysWebpart/components/MonthBirthdaysWebpart.tsx` - Main UI component
- **Data fetching**: Uses SPHttpClient with OData v3 to fetch from SharePoint list at hardcoded endpoint `https://ebyorgar.sharepoint.com/sites/intranet_EBY/_api/web/lists/getbytitle('birthdays_eby')/items`

### Key Technical Details
- **Framework**: SharePoint Framework 1.21.0 with React 17.0.1
- **UI Libraries**: Fluent UI React components (@fluentui/react and @fluentui/react-components)
- **Virtualization**: Uses react-window FixedSizeList for performance with large lists
- **Date handling**: Uses date-fns library for date formatting and manipulation
- **Styling**: SCSS modules with Fluent UI styling

### Data Model
```typescript
interface IMonthBirthday {
  Title: string;    // Full name
  field_1: string;  // Birthday in DD/MM format
}
```

### Business Logic
- Fetches all birthdays from SharePoint list (up to 2000 items)
- Filters to show only current month birthdays from today onwards
- Sorts by day of month
- Highlights today's birthdays with larger persona and special message
- Configurable max display number via property pane (2-5 items)

### Environment Detection
The webpart detects and adapts to different Microsoft 365 environments (SharePoint, Teams, Outlook, Office).

## Development Notes
- The component name inconsistencies exist (TestWebpartNode22 vs MonthBirthdaysWebpart) - this is legacy naming
- Spanish text is used throughout the UI
- Hardcoded SharePoint site URL needs to be made configurable for different environments
- OData v3 configuration is specifically set up for SharePoint REST API compatibility