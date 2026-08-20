import { defineCollection, z } from 'astro:content';

const trackSchema = z.enum(['beginner', 'intermediate', 'advanced']);

const lessons = defineCollection({
  type: 'content',
  schema: z.object({
    title: z.string(),
    track: trackSchema,
    order: z.number(),
    summary: z.string(),
    featured: z.boolean().default(false),
    draft: z.boolean().default(false),
  }),
});

const articles = defineCollection({
  type: 'content',
  schema: z.object({
    title: z.string(),
    summary: z.string(),
    publishedDate: z.coerce.date(),
    tags: z.array(z.string()).default([]),
    draft: z.boolean().default(false),
  }),
});

const courses = defineCollection({
  type: 'content',
  schema: z.object({
    title: z.string(),
    summary: z.string(),
    price: z.string(),
    checkoutUrl: z.string().url(),
    featured: z.boolean().default(false),
    draft: z.boolean().default(false),
    outcomes: z.array(z.string()).default([]),
    includes: z.array(z.string()).default([]),
  }),
});

export const collections = { lessons, articles, courses };

export type Track = z.infer<typeof trackSchema>;
