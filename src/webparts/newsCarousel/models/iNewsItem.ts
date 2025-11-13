export interface INewsItem {
  id: number;
  title: string;
  description: string;
  imageUrl: string;
  pageUrl: string;
  publishedDate: Date;
  author: string;
  category: string;
  isFeatured: boolean;
}