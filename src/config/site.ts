export type Track = 'beginner' | 'intermediate' | 'advanced';

export const site = {
  name: 'Decoding VBA',
  title: 'Decoding VBA | Learn Excel VBA for Workplace Automation',
  description:
    'Free structured lessons and practical guides for Excel VBA. Learn automation for reporting, validation, and month-end workflows.',
  tagline: 'Clear VBA for the Excel work you do every day.',
  url: 'https://decodingvba.com',
  email: 'hello@decodingvba.com',
  nav: [
    { label: 'Learn', href: '/learn/' },
    { label: 'Articles', href: '/articles/' },
    { label: 'Courses', href: '/courses/' },
    { label: 'Search', href: '/search/' },
    { label: 'About', href: '/about/' },
  ],
} as const;

export const analytics = {
  provider: 'Plausible',
  plausibleDomain: 'decodingvba.com',
  policyUrl: 'https://plausible.io/data-policy',
  sitePolicyUrl: 'https://plausible.io/privacy',
} as const;

export const tracks: Record<
  Track,
  { title: string; description: string; href: string }
> = {
  beginner: {
    title: 'Beginner',
    description:
      'Open the VBE, write your first macros, and automate simple worksheet tasks.',
    href: '/learn/beginner/',
  },
  intermediate: {
    title: 'Intermediate',
    description:
      'Handle data with arrays and dictionaries, add error handling, and build reusable tools.',
    href: '/learn/intermediate/',
  },
  advanced: {
    title: 'Advanced',
    description:
      'Design class modules, event-driven workflows, and performant automation apps.',
    href: '/learn/advanced/',
  },
};

export const checkout = {
  provider: 'Lemon Squeezy',
  placeholderUrl: 'https://decodingvba.lemonsqueezy.com/buy/placeholder',
  note: 'Replace placeholderUrl with your Lemon Squeezy or Gumroad checkout link.',
} as const;
