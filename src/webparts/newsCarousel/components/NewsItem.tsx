import * as React from 'react';
import styles from './NewsItem.module.scss';
import { INewsItem } from '../models/iNewsItem';

export type CarouselPosition = 'prev' | 'current' | 'next';

export interface INewsItemProps {
  item: INewsItem;
  position: CarouselPosition;
  onSelect: () => void;
}

export const NewsItem: React.FC<INewsItemProps> = ({ item, position, onSelect }) => {
  const isCurrent = position === 'current';

  const handleClick = (): void => {
    if (isCurrent) {
      window.open(item.pageUrl, '_blank');
    } else {
      onSelect();
    }
  };

  return (
    <article
      className={`${styles.newsItem} ${styles[position]} ${isCurrent ? styles.current : ''}`}
      onClick={handleClick}
      tabIndex={0}
      onKeyDown={(event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          handleClick();
        }
      }}
      role="button"
      aria-label={`${isCurrent ? 'Open' : 'Navigate to'} ${item.title}`}
    >
      <div className={styles.imageContainer}>
        <img src={item.imageUrl} alt={item.title} className={styles.image} />
      </div>
      <div className={styles.content}>
        <h3 className={styles.title}>{item.title}</h3>
        <p className={styles.description}>{item.description}</p>
        <div className={styles.meta}>
          <span className={styles.date}>
            {item.publishedDate.toLocaleDateString()}
          </span>
        </div>
      </div>
    </article>
  );
};