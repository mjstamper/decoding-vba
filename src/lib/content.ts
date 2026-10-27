import type { CollectionEntry } from 'astro:content';
import type { Track } from './types';

export function sortLessons(lessons: CollectionEntry<'lessons'>[]) {
  return [...lessons].sort((a, b) => a.data.order - b.data.order);
}

export function getLessonsByTrack(
  lessons: CollectionEntry<'lessons'>[],
  track: Track,
) {
  return sortLessons(lessons.filter((lesson) => lesson.data.track === track));
}

export function getAdjacentLessons(
  lessons: CollectionEntry<'lessons'>[],
  current: CollectionEntry<'lessons'>,
) {
  const trackLessons = getLessonsByTrack(lessons, current.data.track);
  const index = trackLessons.findIndex((lesson) => lesson.id === current.id);
  return {
    prev: index > 0 ? trackLessons[index - 1] : undefined,
    next: index < trackLessons.length - 1 ? trackLessons[index + 1] : undefined,
  };
}

export function lessonUrl(lesson: CollectionEntry<'lessons'>) {
  return `/learn/${lesson.data.track}/${lesson.id}/`;
}

export function formatDate(date: Date) {
  return new Intl.DateTimeFormat('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }).format(date);
}

export function getAllArticleTags(articles: CollectionEntry<'articles'>[]) {
  const tags = new Set<string>();
  for (const article of articles) {
    if (article.data.draft) continue;
    for (const tag of article.data.tags) {
      tags.add(tag);
    }
  }
  return [...tags].sort((a, b) => a.localeCompare(b));
}

export function tagUrl(tag: string) {
  return `/articles/tags/${encodeURIComponent(tag)}/`;
}

export function getRelatedArticles(
  articles: CollectionEntry<'articles'>[],
  current: CollectionEntry<'articles'>,
  limit = 4,
) {
  const published = articles.filter((a) => !a.data.draft && a.id !== current.id);
  const currentTags = new Set(current.data.tags);

  const scored = published
    .map((article) => ({
      article,
      score: article.data.tags.filter((tag) => currentTags.has(tag)).length,
    }))
    .filter(({ score }) => score > 0)
    .sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      return (
        b.article.data.publishedDate.getTime() - a.article.data.publishedDate.getTime()
      );
    });

  if (scored.length >= limit) {
    return scored.slice(0, limit).map(({ article }) => article);
  }

  const picked = new Set(scored.map(({ article }) => article.id));
  const fallback = published
    .filter((a) => !picked.has(a.id))
    .sort(
      (a, b) => b.data.publishedDate.getTime() - a.data.publishedDate.getTime(),
    );

  return [...scored.map(({ article }) => article), ...fallback].slice(0, limit);
}

export function getRelatedLessons(
  lessons: CollectionEntry<'lessons'>[],
  current: CollectionEntry<'lessons'>,
  limit = 4,
) {
  const trackLessons = getLessonsByTrack(lessons, current.data.track).filter(
    (lesson) => lesson.id !== current.id,
  );

  return trackLessons
    .sort(
      (a, b) =>
        Math.abs(a.data.order - current.data.order) -
        Math.abs(b.data.order - current.data.order),
    )
    .slice(0, limit);
}

export function getArticlesForTrack(
  articles: CollectionEntry<'articles'>[],
  track: Track,
  excludeId?: string,
  limit = 3,
) {
  return articles
    .filter((a) => !a.data.draft && a.id !== excludeId)
    .filter((a) => a.data.tags.includes(track))
    .sort(
      (a, b) => b.data.publishedDate.getTime() - a.data.publishedDate.getTime(),
    )
    .slice(0, limit);
}
