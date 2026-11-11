// The viewer shows no IeumOffice UI text, so the app's translation tables are left out.
export function useI18n() { return { lang: 'ko', t: (key) => key, dateLocale: 'ko-KR' } }
