export const expandSchoolName = (rawName: string): string => {
  const acronyms: Record<string, string> = {
    'MPPS': 'Mandal Parishad Primary School',
    'MPUPS': 'Mandal Parishad Upper Primary School',
    'GPS': 'Government Primary School',
    'GGPS': 'Government Girls Primary School',
    'CPS': 'Primary School run under Central Govt / Cantonment',
    'MPHS': 'Mandal Parishad High School',
  };

  const parts = rawName.trim().split(' ');
  if (parts.length === 0) return rawName;

  const firstWord = parts[0].toUpperCase();
  
  if (acronyms[firstWord]) {
    const fullForm = acronyms[firstWord];
    const villageName = parts.slice(1).join(' ');
    
    // If there is a village name, append it with a comma
    if (villageName) {
      return `${fullForm}, ${villageName}`;
    }
    // If it's just the acronym, return just the full form
    return fullForm;
  }

  // If no matching acronym, return original
  return rawName;
};
