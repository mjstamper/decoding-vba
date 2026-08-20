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
