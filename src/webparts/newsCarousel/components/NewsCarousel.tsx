import * as React from 'react';
import { INewsCarouselProps } from './INewsCarouselProps';
import { NewsService } from '../services/NewsService';
import { NewsItem } from './NewsItem';
import styles from '../components/NewsCarousel.module.scss';
import { INewsItem } from '../models/iNewsItem';

export default class NewsCarousel extends React.Component<INewsCarouselProps, {
  newsItems: INewsItem[];
  currentIndex: number;
  loading: boolean;
  error: string;
  isPaused: boolean;
}> {
  private newsService: NewsService;
  private autoPlayTimer: number = 0;
  private carouselRef: React.RefObject<HTMLDivElement>;
  private containerRef: React.RefObject<HTMLDivElement>;
  private styleObserver: MutationObserver | null = null;

  constructor(props: INewsCarouselProps) {
    super(props);
    this.state = {
      newsItems: [],
      currentIndex: 0,
      loading: true,
      error: '',
      isPaused: false
    };
    this.newsService = new NewsService(this.props.context);
    this.carouselRef = React.createRef<HTMLDivElement>();
    this.containerRef = React.createRef<HTMLDivElement>();
  }

  public async componentDidMount(): Promise<void> {
    await this.loadNewsItems();
    // Always enable auto-scroll
    this.startAutoPlay();
    
    // Inject style into document head for maximum priority
    this.injectGlobalStyles();
    
    // Ensure width is applied after mount (in case SharePoint overrides styles)
    this.applyWidthStyles();
    
    // Use MutationObserver to watch for style changes and reapply
    this.watchForStyleChanges();
  }
  
  private injectGlobalStyles(): void {
    // Check if style already exists
    const existingStyle = document.getElementById('news-carousel-width-override');
    if (existingStyle) {
      existingStyle.remove();
    }
    
    const style = document.createElement('style');
    style.id = 'news-carousel-width-override';
    style.textContent = `
      [class*="newsCarousel_"] {
        width: 100% !important;
      }
    `;
    document.head.appendChild(style);
  }
  
  private applyWidthStyles(): void {
    const applyStyles = (): void => {
      // Only apply 100% width to newsCarousel, not carouselContainer
      if (this.carouselRef.current) {
        this.carouselRef.current.style.setProperty('width', '100%', 'important');
        this.carouselRef.current.style.setProperty('margin', '0', 'important');
      }
      
      // Find and override all newsCarousel elements
      const carousels = document.querySelectorAll('[class*="newsCarousel_"]');
      carousels.forEach((carousel) => {
        const el = carousel as HTMLElement;
        el.style.setProperty('width', '100%', 'important');
        el.style.setProperty('margin', '0', 'important');
      });
    };
    
    // Apply immediately
    applyStyles();
    
    // Apply again after a delay to override any late-loading styles
    setTimeout(applyStyles, 100);
    setTimeout(applyStyles, 500);
    setTimeout(applyStyles, 1000);
  }
  
  private watchForStyleChanges(): void {
    if (!this.carouselRef.current) return;
    
    this.styleObserver = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        if (mutation.type === 'attributes' && mutation.attributeName === 'style') {
          // Style was changed, reapply our styles
          this.applyWidthStyles();
        }
      });
    });
    
    this.styleObserver.observe(this.carouselRef.current, {
      attributes: true,
      attributeFilter: ['style', 'class']
    });
  }

  public componentWillUnmount(): void {
    if (this.autoPlayTimer) {
      clearInterval(this.autoPlayTimer);
      this.autoPlayTimer = 0;
    }
    
    if (this.styleObserver) {
      this.styleObserver.disconnect();
    }
    
    // Remove injected style
    const style = document.getElementById('news-carousel-width-override');
    if (style) {
      style.remove();
    }
  }

  private async loadNewsItems(): Promise<void> {
    try {
      const items = await this.newsService.getNewsItems();
      if (items.length === 0) {
        console.warn('No news items found. Check browser console for ContentType details.');
      }
      // Limit to 9 items for the carousel - IMPORTANT: Only show first 9 items
      const MAX_ITEMS = 9;
      const limitedItems = items.slice(0, MAX_ITEMS);
      console.log(`[News Carousel] Total items fetched: ${items.length}`);
      console.log(`[News Carousel] Limiting to ${MAX_ITEMS} items for carousel display`);
      console.log(`[News Carousel] Items in carousel: ${limitedItems.length}`);
      console.log(`[News Carousel] Item titles:`, limitedItems.map(item => item.title));
      
      this.setState({
        newsItems: limitedItems, // Only these 9 items will be in the carousel
        loading: false,
        error: items.length === 0 ? 'No news articles found. Please check the browser console for details.' : ''
      });
    } catch (error) {
      console.error('Error loading news items:', error);
      this.setState({
        error: `Failed to load news items: ${error instanceof Error ? error.message : 'Unknown error'}`,
        loading: false
      });
    }
  }

  private getViewAllNewsUrl(): string {
    // Construct URL to the SharePoint news page
    const webAbsoluteUrl = this.props.context.pageContext.web.absoluteUrl;
    return `${webAbsoluteUrl}/_layouts/15/news.aspx`;
  }

  private startAutoPlay(): void {
    if (this.autoPlayTimer) {
      clearInterval(this.autoPlayTimer);
    }
    this.autoPlayTimer = window.setInterval(() => {
      this.nextSlide();
    }, this.props.autoPlayInterval);
    this.setState({ isPaused: false });
  }

  private stopAutoPlay(): void {
    if (this.autoPlayTimer) {
      clearInterval(this.autoPlayTimer);
      this.autoPlayTimer = 0;
    }
    this.setState({ isPaused: true });
  }

  private toggleAutoPlay = (): void => {
    if (this.state.isPaused) {
      this.startAutoPlay();
    } else {
      this.stopAutoPlay();
    }
  }

  private nextSlide = (): void => {
    const { newsItems, currentIndex } = this.state;
    const nextIndex = (currentIndex + 1) % newsItems.length;
    this.setState({ currentIndex: nextIndex });
  }

  private prevSlide = (): void => {
    const { newsItems, currentIndex } = this.state;
    const prevIndex = currentIndex === 0 ? newsItems.length - 1 : currentIndex - 1;
    this.setState({ currentIndex: prevIndex });
  }

  private goToSlide = (index: number): void => {
    this.setState({ currentIndex: index });
  }

  public render(): React.ReactElement<INewsCarouselProps> {
    const { newsItems, currentIndex, loading, error } = this.state;
    const { title } = this.props;

    // Debug logging to verify what's being rendered
    if (!loading && newsItems.length > 0) {
      console.log(`[News Carousel Render] Rendering ${newsItems.length} items in carousel`);
      console.log(`[News Carousel Render] Current index: ${currentIndex}`);
      console.log(`[News Carousel Render] Number of dots: ${newsItems.length}`);
    }

    const containerStyle = { width: '100% !important', margin: '0 !important', maxWidth: 'none !important', display: 'block !important' };

    if (loading) {
      return (
        <div ref={this.carouselRef} className={styles.newsCarousel} style={containerStyle}>
          <div className={styles.loading}>Nachrichten werden geladen...</div>
        </div>
      );
    }

    if (error) {
      return (
        <div ref={this.carouselRef} className={styles.newsCarousel} style={containerStyle}>
          <div className={styles.error}>{error}</div>
        </div>
      );
    }

    if (newsItems.length === 0) {
      return (
        <div ref={this.carouselRef} className={styles.newsCarousel} style={containerStyle}>
          <div className={styles.noNews}>Keine Nachrichten gefunden.</div>
        </div>
      );
    }

    type CarouselPosition = 'prev' | 'current' | 'next';

    const computeOffsets = (totalItems: number): number[] => {
      if (totalItems >= 3) {
        return [-1, 0, 1];
      }

      switch (totalItems) {
        case 2:
          return [0, 1];
        default:
          return [0];
      }
    };

    const mapOffsetToPosition = (offset: number): CarouselPosition => {
      switch (offset) {
        case -1:
          return 'prev';
        case 1:
          return 'next';
        default:
          return 'current';
      }
    };

    const getVisibleItems = (): Array<{
      item: INewsItem;
      position: CarouselPosition;
      slideIndex: number;
    }> => {
      const totalItems = newsItems.length;
      const offsets = computeOffsets(totalItems);

      return offsets.map(offset => {
        const index = (currentIndex + offset + totalItems) % totalItems;
        return {
          item: newsItems[index],
          position: mapOffsetToPosition(offset),
          slideIndex: index
        };
      });
    };

    return (
      <div ref={this.carouselRef} className={styles.newsCarousel} style={containerStyle}>
          <div className={(styles as Record<string, string>).header}>
            <h2 className={styles.title}>{title}</h2>
            <a target="_blank" rel="noopener noreferrer" data-interception="off" href={this.getViewAllNewsUrl()} className={(styles as Record<string, string>).viewAllButton}>
              Alle Anzeigen
            </a>
          </div>
          <div ref={this.containerRef} className={styles.carouselContainer}>
          <div className={styles.carousel}>
            {getVisibleItems().map(({ item, position, slideIndex }) => (
              <NewsItem
                key={`${item.id}-${position}-${slideIndex}`}
                item={item}
                position={position}
                onSelect={() => this.goToSlide(slideIndex)}
              />
            ))}
          </div>
        </div>
        
        <div className={styles.navigation}>
          <button className={styles.pauseButton} onClick={this.toggleAutoPlay}>
            <div className={styles.pauseIcon}>
              {this.state.isPaused ? (
                <svg width="12" height="12" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M6 4L12 8L6 12V4Z" fill="#323130"/>
                </svg>
              ) : (
                <svg width="12" height="12" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect x="5" y="4" width="2" height="8" rx="1" fill="#323130"/>
                  <rect x="9" y="4" width="2" height="8" rx="1" fill="#323130"/>
                </svg>
              )}
            </div>
          </button>
          
          <button className={styles.arrowNavLeft} onClick={this.prevSlide}>
            <svg className={styles.arrowNavIcon} width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M12.5 15L7.5 10L12.5 5" stroke="#323130" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </button>
          
          <div className={styles.dots}>
            {newsItems.map((_, index) => (
              <button
                key={index}
                className={`${styles.dot} ${index === currentIndex ? styles.active : ''}`}
                onClick={() => this.goToSlide(index)}
              />
            ))}
          </div>
          
          <button className={styles.arrowNavRight} onClick={this.nextSlide}>
            <svg className={styles.arrowNavIcon} width="20" height="20" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M7.5 5L12.5 10L7.5 15" stroke="#323130" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
            </svg>
          </button>
        </div>
      </div>
    );
  }
}