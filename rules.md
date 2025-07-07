# MedSpa SEO Application Architecture

## Overview
This document outlines the architecture and key components of the MedSpa SEO application, which helps track and analyze SEO performance for medical spa businesses.

## Workbook Architecture
The application uses a two-workbook system:

### 1. Client Workbook
- **Purpose**: Stores individual client data and competitor analysis
- **Structure**: One unique Google Sheet per client
- **Content**:
  - Website statistics
  - Google Business Profile (GBP) information
  - Competitor data (tracks 4 competitors)
  - Geogrid analysis
  - On-Page Insights
- **Script**: Managed by `Client Workbook App Script v2.gs`
- **Access**: Frontend identifies specific workbook via `workbookId` URL parameter

### 2. Services Workbook
- **Purpose**: Manages service cards data for the dashboard
- **Structure**: Single Google Sheet shared across all clients
- **Content**: Data for two service cards displayed on the dashboard
- **Script**: Managed by `Service Cards (Dashboard) App Script.gs`

## Frontend Components

### Dashboard
- Displays GBP Insights with AI-generated SWOT analysis from client workbook
- Features summary metrics and key performance indicators from client workbook
- Shows service cards from Services Workbook (this is the only data which comes from the service workbook)

### Website Stats
- Shows technical SEO metrics
- Displays AI-generated website analysis report from On-Page Insights
- Features website health metrics:
  - Overall Page Score (with gauge graphic)
  - Key Health Metrics (Site Speed, SSL status, Checks Passed)
  - Issues to Review (errors and warnings)

### Backlinks Analysis
- Comprehensive backlink profile analysis
- Shows link types distribution (dofollow/nofollow)
- Tracks link dynamics (new/lost links)
- Competitor backlink comparison

### Geogrid Maps
- Visualizes geographical performance data
- Tracks local search visibility
- Maps competitor presence
- graphs statistics on avg geogrid score and number of hits in top5

## Data Flow
1. Frontend makes requests to Google App Scripts
2. Scripts fetch data from respective workbooks
3. Data is processed and returned to frontend
4. Frontend components render visualizations and metrics

## Key Features
- Real-time data synchronization
- Competitor tracking and analysis
- AI-powered insights and recommendations
- Technical SEO monitoring
- Local SEO performance tracking

## Best Practices
- Use URL parameters to identify specific client workbooks
- Cache responses to minimize API calls
- Handle loading and error states appropriately
- Maintain consistent styling across components

## Technical Notes
- Frontend built with React
- Uses Recharts for data visualization
- Implements responsive design
- Uses Tailwind CSS for styling
- Includes error handling and loading states

## Future Considerations
- Tools section (marked as "coming soon")
- Additional competitor analysis features
- Enhanced AI reporting capabilities 