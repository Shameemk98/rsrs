
export const spNavigate = (url: string) => {
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('data-intercept', 'on');
    link.style.display = 'none';
  
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
  