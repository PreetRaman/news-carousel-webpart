import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INewsItem } from '../models/iNewsItem';

interface IBannerImageValue {
  Url?: string;
  url?: string;
  src?: string;
  value?: string;
}

interface ISharePointListItem {
  Id: number;
  Title: string;
  Description?: string;
  FileRef?: string;
  FileLeafRef?: string;
  Modified: string;
  Author?: {
    Title: string;
  };
  CreatedBy?: {
    Title: string;
  };
  ContentType?: {
    Name: string;
  };
  ContentTypeId?: string;
  PromotedState?: number;
  BannerImageUrl?: string | IBannerImageValue | undefined;
  CanvasContent1?: string;
}

interface IListInfo {
  Id: string;
  Title: string;
  BaseTemplate: number;
  ItemCount: number;
}

interface IListItemWrapper {
  ListItemAllFields?: ISharePointListItem;
  Name?: string;
  ServerRelativeUrl?: string;
  TimeLastModified?: string;
}

type ListIdentifier = { type: 'title' | 'guid'; value: string };
type AugmentedResponse = SPHttpClientResponse & {
  _jsonData?: { value?: unknown[] };
};

export class NewsService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  public async getNewsItems(): Promise<INewsItem[]> {
    try {
      // First, try to find "Neuigkeiten" list directly by querying all lists and searching by name
      let sitePagesListTitle: string | undefined;
      let sitePagesListId: string | undefined;
      
      try {
        const allListsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount`;
        const allListsResponse = await this.context.spHttpClient.get(
          allListsUrl,
          SPHttpClient.configurations.v1
        );
        
        if (allListsResponse.ok) {
          const allListsJson = await allListsResponse.json() as { value?: IListInfo[] };
          const allLists: IListInfo[] = Array.isArray(allListsJson.value) ? allListsJson.value : [];
          
          // Search for "Neuigkeiten" specifically
          const neuigkeitenList = this.findFirst(allLists, (list: IListInfo) => {
            const titleLower = (list.Title || '').toLowerCase();
            return titleLower === 'neuigkeiten' || this.stringIncludes(titleLower, 'neuigkeit');
          });
          
          if (neuigkeitenList && neuigkeitenList.ItemCount > 0) {
            sitePagesListTitle = neuigkeitenList.Title;
            sitePagesListId = neuigkeitenList.Id;
            console.log(`✓ Found "Neuigkeiten" list directly: "${sitePagesListTitle}" (BaseTemplate: ${neuigkeitenList.BaseTemplate}, Items: ${neuigkeitenList.ItemCount})`);
            console.log(`  List ID: ${sitePagesListId}`);
          }
        }
      } catch (err) {
        console.log('Could not query all lists for Neuigkeiten, trying BaseTemplate search...', err);
      }
      
      // If not found by name, try to find the Site Pages list by BaseTemplate
      // Site Pages list can have different BaseTemplate values:
      // - 850 (WebPageLibrary - classic)
      // - 160 (Site Pages - actual pages library)
      // - 3415 (Modern Pages - but this might be template extensions, not actual pages)
      // Try 160 first as it's the actual Site Pages library
      if (!sitePagesListTitle) {
        const baseTemplates = [160, 850, 3415];
        
        for (const baseTemplate of baseTemplates) {
          try {
            const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount&$filter=BaseTemplate eq ${baseTemplate}`;
            const listsResponse = await this.context.spHttpClient.get(
              listsUrl,
        SPHttpClient.configurations.v1
      );

            if (listsResponse.ok) {
              const listsJson = await listsResponse.json() as { value?: IListInfo[] };
              const sitePagesLists: IListInfo[] = Array.isArray(listsJson.value) ? listsJson.value : [];
              
              if (sitePagesLists.length > 0) {
                // Filter out template extension lists - they're not the Site Pages list
                const validLists = sitePagesLists.filter((list: IListInfo) => {
                  const titleLower = (list.Title || '').toLowerCase();
                  // Exclude template extension lists
                  return !this.stringIncludes(titleLower, 'vorlagenerweiterung') && 
                         !this.stringIncludes(titleLower, 'template extension') &&
                         list.ItemCount > 0; // Only lists with items
                });
                
                if (validLists.length > 0) {
                  // Prefer lists with names that suggest they're Site Pages
                  const preferredList = this.findFirst(validLists, (list: IListInfo) => {
                    const titleLower = (list.Title || '').toLowerCase();
                    return this.stringIncludes(titleLower, 'neuigkeit') || // German "News" - highest priority
                           this.stringIncludes(titleLower, 'news') ||
                           this.stringIncludes(titleLower, 'seite') || 
                           this.stringIncludes(titleLower, 'page') ||
                           this.stringIncludes(titleLower, 'nachricht');
                  }) || validLists[0];
                  
                  sitePagesListTitle = preferredList.Title;
                  sitePagesListId = preferredList.Id;
                  console.log(`✓ Found Site Pages list: "${sitePagesListTitle}" (BaseTemplate: ${baseTemplate}, Items: ${preferredList.ItemCount})`);
                  console.log(`  List ID: ${sitePagesListId}`);
                  console.log(`  All lists with BaseTemplate ${baseTemplate}:`, sitePagesLists.map((l: IListInfo) => ({ Title: l.Title, ItemCount: l.ItemCount })));
                  break;
                } else {
                  console.log(`  BaseTemplate ${baseTemplate} lists found but none are valid Site Pages lists (filtered out template extensions)`);
                }
              }
            }
          } catch (err) {
            console.log(`Error checking BaseTemplate ${baseTemplate}:`, err);
            // Try next BaseTemplate
            continue;
          }
        }
      }
      
      // If BaseTemplate filter didn't work, try getting all lists and filtering manually
      if (!sitePagesListTitle) {
        console.warn('No Site Pages list found with BaseTemplate filter. Trying to get all lists and filter manually...');
        try {
          const allListsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount`;
          const allListsResponse = await this.context.spHttpClient.get(
            allListsUrl,
            SPHttpClient.configurations.v1
          );
          
          if (allListsResponse.ok) {
            const allListsJson = await allListsResponse.json() as { value?: IListInfo[] };
            const allLists: IListInfo[] = Array.isArray(allListsJson.value) ? allListsJson.value : [];
            
            // Manually filter for Site Pages lists, excluding template extensions
            const sitePagesLists = allLists.filter((list: IListInfo) => {
              const titleLower = (list.Title || '').toLowerCase();
              const isCorrectBaseTemplate = list.BaseTemplate === 850 || 
                                          list.BaseTemplate === 3415 || 
                                          list.BaseTemplate === 160;
              const isNotTemplateExtension = !this.stringIncludes(titleLower, 'vorlagenerweiterung') && 
                                           !this.stringIncludes(titleLower, 'template extension');
              return isCorrectBaseTemplate && isNotTemplateExtension && list.ItemCount > 0;
            });
            
            if (sitePagesLists.length > 0) {
              // Prefer lists with names that suggest they're Site Pages
              const preferredList = this.findFirst(sitePagesLists, (list: IListInfo) => {
                const titleLower = (list.Title || '').toLowerCase();
                return this.stringIncludes(titleLower, 'neuigkeit') ||
                       this.stringIncludes(titleLower, 'news') ||
                       this.stringIncludes(titleLower, 'seite') || 
                       this.stringIncludes(titleLower, 'page') ||
                       this.stringIncludes(titleLower, 'nachricht');
              }) || sitePagesLists[0];
              
              sitePagesListTitle = preferredList.Title;
              sitePagesListId = preferredList.Id;
              console.log(`✓ Found Site Pages list manually: "${sitePagesListTitle}" (BaseTemplate: ${preferredList.BaseTemplate}, Items: ${preferredList.ItemCount})`);
              console.log(`  List ID: ${sitePagesListId}`);
              
              // Also log all potential lists for debugging
              console.log('  All potential Site Pages lists:', sitePagesLists.map((l: IListInfo) => ({ Title: l.Title, BaseTemplate: l.BaseTemplate, ItemCount: l.ItemCount })));
            } else {
              console.warn('No Site Pages list found even after getting all lists.');
              console.log('  Available lists with BaseTemplate 160/850/3415:', allLists.filter((l: IListInfo) => l.BaseTemplate === 160 || l.BaseTemplate === 850 || l.BaseTemplate === 3415).map((l: IListInfo) => ({ Title: l.Title, BaseTemplate: l.BaseTemplate, ItemCount: l.ItemCount })));
            }
          }
        } catch (err) {
          console.warn('Could not get all lists, trying direct list names...', err);
        }
      }

      // Try multiple list names and methods
      // Also try accessing by GUID if we have the list ID
      const listIdentifiers: ListIdentifier[] = [];
      
      // First try by GUID if we have it (most reliable)
      if (sitePagesListId) {
        listIdentifiers.push({ type: 'guid', value: sitePagesListId });
      }
      
      // Then try by title - prioritize "Neuigkeiten"
      if (sitePagesListTitle) {
        listIdentifiers.push({ type: 'title', value: sitePagesListTitle });
      }
      
      // Add common list name variations
      listIdentifiers.push(
        { type: 'title', value: 'Neuigkeiten' },      // Site Pages list name (HIGHEST PRIORITY)
        { type: 'title', value: 'SitePages' },        // Internal name (language-independent)
        { type: 'title', value: 'Site Pages' },       
        { type: 'title', value: 'Seiten' },           
        { type: 'title', value: 'Webseiten' },        
        { type: 'title', value: 'Modern Pages' },     
        { type: 'title', value: 'Moderne Seiten' }    
      );
      
      // Also try accessing Site Pages directly via server-relative URL (works even if hidden)
      // This is a common workaround for permission issues - try this FIRST before other methods
      const sitePagesServerRelativeUrl = `${this.context.pageContext.web.serverRelativeUrl}/SitePages`;
      // Insert at the beginning to try this first
      listIdentifiers.unshift({ type: 'guid', value: sitePagesServerRelativeUrl });
      
      let response: AugmentedResponse | undefined;
      let successfulListName: string | undefined;
      
      // Try each list identifier until one works
      for (const listInfo of listIdentifiers) {
        try {
          // Query without Description field first (it may not exist in all lists)
          // Don't use $expand=Author as it can cause permission errors - use CreatedBy instead
          // Try with BannerImageUrl first, but fallback if field doesn't exist
          let allPagesUrl: string;
          let allPagesUrlFallback: string;
          
          if (listInfo.type === 'guid') {
            // Check if it's a GUID or a server-relative URL
            if (listInfo.value.indexOf('/') >= 0) {
              // It's a server-relative URL - use GetFolderByServerRelativeUrl
              allPagesUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(listInfo.value)}')/Files?$select=ListItemAllFields/Id,ListItemAllFields/Title,ListItemAllFields/FileRef,ListItemAllFields/FileLeafRef,ListItemAllFields/Modified,ListItemAllFields/CreatedBy/Title,ListItemAllFields/ContentType/Name,ListItemAllFields/ContentTypeId,ListItemAllFields/BannerImageUrl,ListItemAllFields/PromotedState&$expand=ListItemAllFields/CreatedBy,ListItemAllFields/ContentType&$orderby=ListItemAllFields/Modified desc&$top=50`;
              allPagesUrlFallback = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(listInfo.value)}')/Files?$select=ListItemAllFields/Id,ListItemAllFields/Title,ListItemAllFields/FileRef,ListItemAllFields/FileLeafRef,ListItemAllFields/Modified,ListItemAllFields/CreatedBy/Title,ListItemAllFields/ContentType/Name,ListItemAllFields/ContentTypeId&$expand=ListItemAllFields/CreatedBy,ListItemAllFields/ContentType&$orderby=ListItemAllFields/Modified desc&$top=50`;
            } else {
              // It's a GUID - use CreatedBy instead of Author to avoid expand issues
              allPagesUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listInfo.value}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId,BannerImageUrl,PromotedState&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
              allPagesUrlFallback = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listInfo.value}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
            }
          } else {
            allPagesUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listInfo.value)}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId,BannerImageUrl,PromotedState&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
            allPagesUrlFallback = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listInfo.value)}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
          }
          
          console.log(`Trying to access list: "${listInfo.value}" (${listInfo.type})`);
          let listResponse = await this.context.spHttpClient.get(
            allPagesUrl,
            SPHttpClient.configurations.v1
          ) as AugmentedResponse;
          response = listResponse;
          
          // If query fails with 400 (likely due to missing fields), try without BannerImageUrl/PromotedState
          if (!response.ok && response.status === 400) {
            console.log(`Query failed with BannerImageUrl, trying without optional fields...`);
            listResponse = await this.context.spHttpClient.get(
              allPagesUrlFallback,
              SPHttpClient.configurations.v1
            ) as AugmentedResponse;
            response = listResponse;
          }

          if (response.ok) {
            // Read and validate the response before accepting it
            const tempResponseJson = await response.json() as { value?: unknown[] };
            const tempItems = this.ensureArray<unknown>(tempResponseJson.value);
            
            // Quick validation: skip if this is clearly not a Site Pages list
            if (tempItems.length > 0) {
              const previewItem = this.toSharePointListItem(tempItems[0]);
              
              const contentType = (previewItem?.ContentType?.Name || '').toLowerCase();
              const title = (previewItem?.Title || '').toLowerCase();
              
              // Check if this looks like an Access Requests list
              if (this.stringIncludes(contentType, 'zugriffsanforderung') || 
                  this.stringIncludes(contentType, 'access request') ||
                  (this.stringIncludes(title, 'für') && this.stringIncludes(title, 'von'))) {
                console.log(`List '${listInfo.value}' appears to be Access Requests, not Site Pages. Skipping...`);
                response = undefined;
                continue;
              }
            }
            
            // This looks like a valid Site Pages list
            successfulListName = listInfo.value;
            // Store the response data to use later
            response._jsonData = tempResponseJson;
            console.log(`Successfully connected to list: "${listInfo.value}"`);
            break;
          } else {
            const errorText = await response.text();
            console.log(`✗ List '${listInfo.value}' not found (${response.status}): ${errorText.substring(0, 150)}`);
            response = undefined;
            continue;
          }
        } catch (err: unknown) {
          console.log(`Error trying list '${listInfo.value}':`, this.getErrorMessage(err));
          response = undefined;
          continue;
        }
      }

      if (!response || !response.ok || !successfulListName) {
        // Last attempt: try to get all lists and show what's available, especially BaseTemplate 160
        try {
          const allListsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount`;
          const allListsResponse = await this.context.spHttpClient.get(
            allListsUrl,
            SPHttpClient.configurations.v1
          );
          if (allListsResponse.ok) {
            const allListsJson = await allListsResponse.json() as { value?: IListInfo[] };
            const allLists = this.ensureArray<IListInfo>(allListsJson.value);
            console.log('Available lists in this site:', allLists);
            
            // Find and log BaseTemplate 160 lists specifically
            const sitePagesLists = allLists.filter((list: IListInfo) => list.BaseTemplate === 160);
            if (sitePagesLists.length > 0) {
              console.log('Lists with BaseTemplate 160 (Site Pages):', sitePagesLists.map((l: IListInfo) => ({ Id: l.Id, Title: l.Title, ItemCount: l.ItemCount })));
              
              // Try to access the first BaseTemplate 160 list by GUID
              const firstSitePagesList = sitePagesLists[0];
              console.log(`Attempting to access BaseTemplate 160 list by GUID: ${firstSitePagesList.Id} (Title: "${firstSitePagesList.Title}")`);
              
              try {
                const guidUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${firstSitePagesList.Id}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId,BannerImageUrl,PromotedState&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
                const guidUrlFallback = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${firstSitePagesList.Id}')/items?$select=Id,Title,FileRef,FileLeafRef,Modified,CreatedBy/Title,ContentType/Name,ContentTypeId&$expand=CreatedBy,ContentType&$orderby=Modified desc&$top=50`;
                
                let guidResponse = await this.context.spHttpClient.get(guidUrl, SPHttpClient.configurations.v1) as AugmentedResponse;
                
                // If query fails with 400, try without optional fields
                if (!guidResponse.ok && guidResponse.status === 400) {
                  console.log(`  Query failed with BannerImageUrl, trying without optional fields...`);
                  guidResponse = await this.context.spHttpClient.get(guidUrlFallback, SPHttpClient.configurations.v1) as AugmentedResponse;
                }
                
                if (guidResponse.ok) {
                  const guidJson = await guidResponse.json() as { value?: unknown[] };
                  const guidItems = this.ensureArray<unknown>(guidJson.value);
                  
                  if (guidItems.length > 0) {
                    // Check if this looks valid
                    const firstItem = this.toSharePointListItem(guidItems[0]);
                    const contentType = (firstItem?.ContentType?.Name || '').toLowerCase();
                    
                    if (!this.stringIncludes(contentType, 'zugriffsanforderung') && !this.stringIncludes(contentType, 'access request')) {
                      console.log(`Successfully accessed list by GUID: "${firstSitePagesList.Title}" with ${guidItems.length} items`);
                      response = guidResponse;
                      successfulListName = firstSitePagesList.Id;
                      response._jsonData = guidJson;
                    }
                  }
                }
              } catch (guidErr) {
                console.log('Could not access list by GUID:', this.getErrorMessage(guidErr));
              }
            }
          }
        } catch (e) {
          console.log(this.getErrorMessage(e));
        }
        
        if (!response || !response.ok || !successfulListName) {
          const errorText = 'Could not find Site Pages list. Please check browser console for available lists.';
          console.error(`Error fetching news items: ${errorText}`);
          throw new Error(errorText);
        }
      }

      // Use the pre-read JSON data if available, otherwise read it now
      const responseJson = response._jsonData ?? await response.json() as { value?: unknown[] };
      const rawItems = this.ensureArray<unknown>(responseJson.value);
      const allItems = rawItems
        .map(raw => this.toSharePointListItem(raw))
        .filter((item): item is ISharePointListItem => item !== undefined);

      console.log(`Total items found in list: ${allItems.length}`);

      // Log detailed information about each item for debugging
      if (allItems.length > 0) {
        const contentTypeNames = allItems
          .map(item => item.ContentType?.Name)
          .filter((name): name is string => Boolean(name));
        const uniqueContentTypes: string[] = [];
        contentTypeNames.forEach((name: string) => {
          if (uniqueContentTypes.indexOf(name) === -1) {
            uniqueContentTypes.push(name);
          }
        });
        console.log('Available ContentTypes:', uniqueContentTypes);
        
        // Log first few items with details
        console.log('Sample items (first 5):');
        allItems.slice(0, 5).forEach((item, index: number) => {
          console.log(`  ${index + 1}. Title: "${item.Title || '(no title)'}" | ContentType: "${item.ContentType?.Name || 'unknown'}" | ContentTypeId: "${item.ContentTypeId || 'unknown'}"`);
        });
      }

      // Filter for article pages (support both English and German ContentTypes)
      // Exclude home pages - filter out pages with titles like "Home", "Startseite" (German), etc.
      // ContentType IDs are language-independent, so we check those first
      let articleItems = allItems.filter((item: ISharePointListItem) => {
        const title = (item.Title || '').toLowerCase();
        const fileLeafRef = (item.FileLeafRef || '').toLowerCase();
        const contentTypeName = (item.ContentType?.Name || '').toLowerCase();
        const contentTypeId = item.ContentTypeId || '';
        
        // Exclude home pages - check title and filename
        if (title === 'home' || title === 'startseite' || title === 'homepage' || 
            fileLeafRef === 'home.aspx' || fileLeafRef === 'startseite.aspx' ||
            fileLeafRef === 'homepage.aspx' || this.stringStartsWith(fileLeafRef, 'home.') ||
            this.stringStartsWith(fileLeafRef, 'startseite.')) {
          console.log(`  Excluding home page: "${item.Title}" (${fileLeafRef})`);
          return false;
        }
        
        // Check ContentType ID (language-independent) - Article Page ContentType ID pattern
        if (this.stringStartsWith(contentTypeId, '0x0101009D1CB255DA76424F860D91F20E6C4118')) {
          return true;
        }
        
        // Check English ContentType names
        if (this.stringIncludes(contentTypeName, 'article') || this.stringIncludes(contentTypeName, 'news')) {
          return true;
        }
        
        // Check German ContentType names
        if (this.stringIncludes(contentTypeName, 'artikel') || this.stringIncludes(contentTypeName, 'artikelseite') || 
            this.stringIncludes(contentTypeName, 'nachricht') || this.stringIncludes(contentTypeName, 'neuigkeit')) {
          return true;
        }
        
        return false;
      });

      console.log(`Items after ContentType filter: ${articleItems.length}`);
      if (articleItems.length === 0 && allItems.length > 0) {
        console.log(`No articles passed ContentType filter. All items:`, allItems.map((item: ISharePointListItem) => ({
          Title: item.Title,
          FileLeafRef: item.FileLeafRef,
          ContentType: item.ContentType?.Name,
          ExcludedAsHome: (item.Title || '').toLowerCase() === 'home' || (item.FileLeafRef || '').toLowerCase() === 'home.aspx'
        })));
      }

      // If no articles found with filter, use all pages (even without title, use FileLeafRef)
      // But still exclude home pages
      if (articleItems.length === 0 && allItems.length > 0) {
        console.warn('No articles found with ContentType filter. Using all pages as fallback.');
        articleItems = allItems.filter((item: ISharePointListItem) => {
          const title = (item.Title || '').toLowerCase();
          const fileLeafRef = (item.FileLeafRef || '').toLowerCase();
          
          // Exclude home pages
          if (title === 'home' || title === 'startseite' || title === 'homepage' || 
              fileLeafRef === 'home.aspx' || fileLeafRef === 'startseite.aspx' ||
              fileLeafRef === 'homepage.aspx' || this.stringStartsWith(fileLeafRef, 'home.') ||
              this.stringStartsWith(fileLeafRef, 'startseite.')) {
            console.log(`  Excluding home page: "${item.Title}" (${fileLeafRef})`);
            return false;
          }
          
          // Use FileLeafRef or FileRef as title if Title is missing
          const hasTitle = item.Title && item.Title.trim() !== '';
          const hasFileRef = item.FileRef || item.FileLeafRef;
          if (!hasTitle && !hasFileRef) {
            console.log(`  Skipping item (no title or file reference): ID ${item.Id}`);
            return false;
          }
          // If no title, use FileLeafRef as title
          if (!hasTitle && hasFileRef) {
            if (item.FileLeafRef) {
              item.Title = item.FileLeafRef;
            } else if (item.FileRef && typeof item.FileRef === 'string') {
              const parts = item.FileRef.split('/');
              item.Title = parts[parts.length - 1] || `Item ${item.Id}`;
            } else {
              item.Title = `Item ${item.Id}`;
            }
            console.log(`  Using FileLeafRef as title for item ${item.Id}: "${item.Title}"`);
          }
          return true;
        });
        console.log(`Using ${articleItems.length} items as news articles (all pages, excluding home)`);
      }

      // Exclude specific news items by title or filename (case-insensitive, partial match)
      articleItems = articleItems.filter((item: ISharePointListItem) => {
        const title = (item.Title || '').trim().toLowerCase();
        const fileLeafRef = (item.FileLeafRef || '').trim().toLowerCase();
        
        // Exclude any news item containing "expo real" (case-insensitive)
        if (this.stringIncludes(title, 'expo real')) {
          console.log(`  Excluding news item: "${item.Title}"`);
          return false;
        }
        
        // Exclude "volltextsuche.aspx" by title or filename
        if (this.stringIncludes(title, 'volltextsuche') || this.stringIncludes(fileLeafRef, 'volltextsuche')) {
          console.log(`  Excluding news item: "${item.Title}" (${item.FileLeafRef})`);
          return false;
        }
        
        // Exclude "Die Spezialisten für Immobilien-Maklerdienste" (case-insensitive)
        if (this.stringIncludes(title, 'die spezialisten für immobilien-maklerdienste') || 
            this.stringIncludes(title, 'spezialisten für immobilien-maklerdienste')) {
          console.log(`  Excluding news item: "${item.Title}"`);
          return false;
        }
        
        return true;
      });

      // Show all available news items (no limit)

      console.log(`Final result: ${articleItems.length} news articles to display`);
      if (articleItems.length > 0) {
        console.log('Article titles:', articleItems.map(item => item.Title));
      }
      
      return this.mapToNewsItems(articleItems);
    } catch (error) {
      console.error('Error fetching news items:', error);
      // Return empty array to show "no news" message
      return [];
    }
  }

  private mapToNewsItems(items: ISharePointListItem[]): INewsItem[] {
    return items.map((item: ISharePointListItem) => {
      // Construct full URL for page in the format: {webAbsoluteUrl}/SitePages/{FileLeafRef}
      // FileLeafRef should be the filename (e.g., "Jobb%C3%B6rse-ae-group-in-Gerstungen.aspx")
      let pageUrl: string;
      
      if (item.FileLeafRef) {
        // Use FileLeafRef directly - it should already be URL encoded
        // Format: /SitePages/filename.aspx
        const fileName = item.FileLeafRef;
        pageUrl = `${this.context.pageContext.web.absoluteUrl}/SitePages/${fileName}`;
      } else if (item.FileRef) {
        // If FileRef is available, extract the filename from it
        // FileRef format: /sites/myWorkSpace/SitePages/filename.aspx
        const fileRef = item.FileRef.indexOf('http') === 0 
          ? new URL(item.FileRef).pathname 
          : item.FileRef;
        const fileName = fileRef.split('/').pop() || `item-${item.Id}.aspx`;
        pageUrl = `${this.context.pageContext.web.absoluteUrl}/SitePages/${fileName}`;
      } else {
        // Fallback
        pageUrl = `${this.context.pageContext.web.absoluteUrl}/SitePages/item-${item.Id}.aspx`;
      }
      
      return {
        id: item.Id,
        title: item.Title || 'Untitled',
        description: this.extractDescription(item.Description, item.CanvasContent1),
        imageUrl: this.getImageUrl(item),
        pageUrl: pageUrl,
        publishedDate: item.Modified ? new Date(item.Modified) : new Date(),
        author: (item.CreatedBy?.Title || item.Author?.Title) || 'Unknown', // Use CreatedBy first, fallback to Author
        category: item.ContentType?.Name || 'Page',
      isFeatured: false
      };
    });
  }

  private extractDescription(description?: string, canvasContent?: string): string {
    const primarySource = (description || '').trim();
    let plainText = primarySource ? this.stripHtml(primarySource) : '';

    if (!plainText && canvasContent) {
      plainText = this.extractTextFromCanvasContent(canvasContent);
    }

    if (!plainText) {
      return '';
    }

    return plainText.length > 200 ? `${plainText.substring(0, 200)}...` : plainText;
  }

  private getImageUrl(item: ISharePointListItem): string {
    // Try to use BannerImageUrl if available (for news articles)
    // BannerImageUrl might be a string, object, or null/undefined, so check type first
    const bannerCandidate = this.normalizeBannerUrl(item.BannerImageUrl);
    if (bannerCandidate) {
      return bannerCandidate;
    }
    
    const canvasImage = this.extractImageFromCanvasContent(item.CanvasContent1);
    if (canvasImage) {
      return canvasImage;
    }

    // Try to get thumbnail from page URL
    // SharePoint pages have a thumbnail endpoint
    if (item.FileLeafRef) {
      // Use SharePoint's thumbnail service for pages
      const previewUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(`${this.context.pageContext.web.serverRelativeUrl}/SitePages/${item.FileLeafRef}`)}&resolution=6`;
      return this.ensurePreviewResolution(previewUrl);
    }
    
    // Fallback to default thumbnail
    return `${this.context.pageContext.web.absoluteUrl}/_layouts/15/images/sitepagethumbnail.png`;
  }

  private normalizeBannerUrl(banner: string | IBannerImageValue | null | undefined): string | null {
    if (!banner) {
      return null;
    }

    if (typeof banner === 'string') {
      const trimmed = banner.trim();
      if (!trimmed) {
        return null;
      }

      // Some APIs return BannerImageUrl as a JSON string
      if (trimmed.charAt(0) === '{' && trimmed.charAt(trimmed.length - 1) === '}') {
        try {
          const parsed = JSON.parse(trimmed);
          return this.normalizeBannerUrl(parsed);
        } catch (error) {
          console.warn('Unable to parse BannerImageUrl JSON string:', error);
        }
      }

      if (trimmed.indexOf('http') === 0) {
        return this.ensurePreviewResolution(trimmed);
      }

      if (trimmed.indexOf('/') === 0) {
        return this.ensurePreviewResolution(`${this.context.pageContext.web.absoluteUrl}${trimmed}`);
      }

      return this.ensurePreviewResolution(`${this.context.pageContext.web.absoluteUrl}/${trimmed}`);
    }

    const candidate = banner.Url || banner.url || banner.src || banner.value;
    if (candidate) {
      return this.normalizeBannerUrl(candidate);
    }

    return null;
  }

  private extractImageFromCanvasContent(canvasContent: string | undefined): string | null {
    if (!canvasContent) {
      return null;
    }

    // Look for data-sp-imgsrc (SharePoint stores image web parts like this)
    const dataSpImgSrcMatch = canvasContent.match(/data-sp-imgsrc="([^"]+)"/i);
    if (dataSpImgSrcMatch && dataSpImgSrcMatch[1]) {
      return this.ensurePreviewResolution(dataSpImgSrcMatch[1]);
    }

    // Fall back to first <img> tag
    const imgSrcMatch = canvasContent.match(/<img[^>]+src="([^"]+)"/i);
    if (imgSrcMatch && imgSrcMatch[1]) {
      return this.ensurePreviewResolution(imgSrcMatch[1]);
    }

    return null;
  }

  private extractTextFromCanvasContent(canvasContent: string): string {
    const cleaned = canvasContent
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<[^>]+>/g, ' ');

    return this.stripHtml(cleaned);
  }

  private stripHtml(value: string): string {
    if (!value) {
      return '';
    }

    const withoutEntities = value
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/gi, "'")
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>');

    return withoutEntities.replace(/\s+/g, ' ').trim();
  }

  private ensureArray<T>(value: unknown): T[] {
    return Array.isArray(value) ? (value as T[]) : [];
  }

  private isSharePointListItem(item: unknown): item is ISharePointListItem {
    return typeof item === 'object' && item !== null && 'Id' in item;
  }

  private isListItemWrapper(item: unknown): item is IListItemWrapper {
    return typeof item === 'object' && item !== null && 'ListItemAllFields' in item;
  }

  private toSharePointListItem(raw: unknown): ISharePointListItem | undefined {
    if (this.isSharePointListItem(raw)) {
      return raw;
    }

    if (this.isListItemWrapper(raw) && raw.ListItemAllFields) {
      const listItem = raw.ListItemAllFields;
      return {
        Id: listItem.Id,
        Title: listItem.Title || raw.Name || 'Untitled',
        FileRef: listItem.FileRef || raw.ServerRelativeUrl,
        FileLeafRef: listItem.FileLeafRef || raw.Name,
        Modified: listItem.Modified || raw.TimeLastModified || new Date().toISOString(),
        CreatedBy: listItem.CreatedBy || listItem.Author,
        Author: listItem.CreatedBy || listItem.Author,
        ContentType: listItem.ContentType,
        ContentTypeId: listItem.ContentTypeId,
        BannerImageUrl: listItem.BannerImageUrl,
        PromotedState: listItem.PromotedState,
        CanvasContent1: listItem.CanvasContent1,
        Description: listItem.Description
      };
    }

    return undefined;
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }

    if (typeof error === 'string') {
      return error;
    }

    return JSON.stringify(error);
  }

  private findFirst<T>(items: T[], predicate: (item: T) => boolean): T | undefined {
    for (const item of items) {
      if (predicate(item)) {
        return item;
      }
    }
    return undefined;
  }

  private stringIncludes(haystack: string, needle: string): boolean {
    return haystack.indexOf(needle) !== -1;
  }

  private stringStartsWith(value: string, prefix: string): boolean {
    return value.indexOf(prefix) === 0;
  }

  private ensurePreviewResolution(url: string | null | undefined): string {
    if (!url) {
      return '';
    }
    const trimmed = url.trim();
    if (!trimmed) {
      return '';
    }

    const lower = trimmed.toLowerCase();
    if (lower.indexOf('getpreview.ashx') === -1) {
      return trimmed;
    }

    if (/[?&]resolution=\d+/i.test(trimmed)) {
      return trimmed;
    }

    const separator = trimmed.indexOf('?') !== -1 ? '&' : '?';
    return `${trimmed}${separator}resolution=6`;
  }
}