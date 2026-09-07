import { LISTS } from "../common/constants";

export const setCache = (key: string, data: any, ttlMinutes = 10) => {
    const payload = {
      data,
      expiry: Date.now() + ttlMinutes * 60 * 1000
    };
  
    sessionStorage.setItem(key, JSON.stringify(payload));
  };
  
  export const getCache = (key: string) => {
    const raw = sessionStorage.getItem(key);
    if (!raw) return null;
  
    try {
      const parsed = JSON.parse(raw);
  
      if (!parsed.expiry || Date.now() > parsed.expiry) {
        sessionStorage.removeItem(key);
        return null;
      }
  
      return parsed.data;
    } catch {
      sessionStorage.removeItem(key);
      return null;
    }
  };
  
  export const clearCache = (key: string) => {
    sessionStorage.removeItem(key);
  };
  
  export  const validateCache = async (
        spService: any,
        listName: string,
        cacheKey: string,
        siteUrl:string
      ): Promise<boolean> => {
      
        const latest = await spService.getLatestModified(listName,siteUrl);
        const cached = sessionStorage.getItem(cacheKey);
      
        if (!cached || cached !== latest) {
          sessionStorage.setItem(cacheKey, latest);
          return false; 
        }
      
        return true; 
      };
  