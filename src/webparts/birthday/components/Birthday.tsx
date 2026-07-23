import * as React from 'react';
import { useState, useEffect } from 'react';
import { IBirthdayProps } from './IBirthdayProps';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './Birthday.module.scss';

interface IBirthdayUser {
  Name: string;
  Birthday: string; 
  UserEmail: string;
}

interface IGroupedBirthdays {
  [key: string]: IBirthdayUser[];
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

export default function Birthday(props: IBirthdayProps): React.ReactElement {
  const [groupedBirthdays, setGroupedBirthdays] = useState<IGroupedBirthdays>({});
  const [expandedMonths, setExpandedMonths] = useState<{ [key: string]: boolean }>({});
  const [loading, setLoading] = useState<boolean>(true);
  const [apiError, setApiError] = useState<string | null>(null);

  useEffect(() => {
    // Automatically expand the current calendar month by default
    const currentMonthName = MONTHS[new Date().getMonth()];
    setExpandedMonths({ [currentMonthName]: true });
    
    fetchBirthdays();
  }, []);

  const fetchBirthdays = async (): Promise<void> => {
    try {
      setApiError(null);
      const subsiteUrl = "https://tns05.sharepoint.com/sites/ProjectManagement";
      const listUrl = `${subsiteUrl}/_api/web/lists/GetByTitle('Birthdays')/items?$select=Birthday,Name/Title,Name/EMail&$expand=Name`;
      
      const response = await props.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
      
      if (!response.ok) {
        throw new Error(`SharePoint REST API Error (Status ${response.status}).`);
      }

      const data = await response.json();

      if (data.value && data.value.length > 0) {
        const processedData: IBirthdayUser[] = data.value
          .filter((item: any) => item.Name) 
          .map((item: any) => ({
            Name: item.Name.Title || "Unknown User",
            Birthday: item.Birthday ? item.Birthday.toString() : "", 
            UserEmail: item.Name.EMail || ""
          }));

        groupUsersByMonth(processedData);
      }
      setLoading(false);
    } catch (error) {
      setApiError(error instanceof Error ? error.message : "An unexpected API connection error occurred.");
      setLoading(false);
    }
  };

  const groupUsersByMonth = (users: IBirthdayUser[]): void => {
    const initialGroups: IGroupedBirthdays = {};
    MONTHS.forEach(month => { initialGroups[month] = []; });

    users.forEach(user => {
      if (!user.Birthday) return;
      let monthIndex = -1;

      const dateObj = new Date(user.Birthday);
      if (!isNaN(dateObj.getTime())) {
        monthIndex = dateObj.getMonth();
      } else {
        const lowerBirthday = user.Birthday.toLowerCase();
        monthIndex = MONTHS.findIndex(m => lowerBirthday.includes(m.toLowerCase()));
      }

      if (monthIndex >= 0 && monthIndex < 12) {
        const monthName = MONTHS[monthIndex];
        initialGroups[monthName].push(user);
      }
    });

    // Sort internal users inside each month by calendar day numerical order
    Object.keys(initialGroups).forEach(month => {
      initialGroups[month].sort((a, b) => {
        const dayA = new Date(a.Birthday).getDate() || 0;
        const dayB = new Date(b.Birthday).getDate() || 0;
        return dayA - dayB;
      });
    });

    setGroupedBirthdays(initialGroups);
  };

  const toggleMonth = (month: string): void => {
    setExpandedMonths(prev => ({ ...prev, [month]: !prev[month] }));
  };

  if (loading) {
    return (
      <div className={styles.birthdayWebPart}>
        <div className={styles.loadingContainer}>
          <div className={styles.spinner}></div>
          <p>Loading celebration calendar...</p>
        </div>
      </div>
    );
  }

  if (apiError) {
    return (
      <div style={{ padding: '24px', maxWidth: '850px', margin: '15px auto' }}>
        <div style={{ 
          padding: '16px', 
          background: '#fde7e9', 
          borderLeft: '4px solid #a80000',
          borderRadius: '4px',
          color: '#a80000',
          fontSize: '14px'
        }}>
          <strong style={{ display: 'block', marginBottom: '4px' }}>Data Connection Issue</strong>
          {apiError}
        </div>
      </div>
    );
  }

  return (
    <div className={styles.birthdayWebPart}>
      <div className={styles.headerSection}>
        <div className={styles.headerIcon}>🎂</div>
        <div>
          <h2 className={styles.title}>Company Birthdays</h2>
          <p className={styles.subtitle}>Celebrate and connect with your colleagues on their special day</p>
        </div>
      </div>
      
      {MONTHS.map(month => {
        const usersInMonth = groupedBirthdays[month] || [];
        const isExpanded = !!expandedMonths[month];
        const isCurrentMonth = MONTHS[new Date().getMonth()] === month;

        return (
          <div 
            key={month} 
            className={`${styles.accordionSection} ${isExpanded ? styles.sectionExpanded : ''} ${isCurrentMonth ? styles.currentMonthHighlight : ''}`}
          >
            <button 
              className={`${styles.accordionHeader} ${isExpanded ? styles.active : ''}`}
              onClick={() => toggleMonth(month)}
            >
              <div className={styles.headerLeft}>
                <span className={styles.monthText}>{month}</span>
                {isCurrentMonth && <span className={styles.currentBadge}>Current Month</span>}
                {usersInMonth.length > 0 && (
                  <span className={styles.countBadge}>{usersInMonth.length}</span>
                )}
              </div>
              <span className={styles.iconArrow}>{isExpanded ? '⚡' : '✨'}</span>
            </button>
            
            {isExpanded && (
              <div className={styles.accordionContent}>
                {usersInMonth.length === 0 ? (
                  <div className={styles.noBirthdaysWrapper}>
                    <p className={styles.noBirthdays}>No birthdays listed for this month.</p>
                  </div>
                ) : (
                  <div className={styles.userGrid}>
                    {usersInMonth.map((user, idx) => {
                      let dayDisplay = "—";
                      const dateObj = new Date(user.Birthday);
                      if (!isNaN(dateObj.getTime())) {
                        const day = dateObj.getDate();
                        const suffix = ["th", "st", "nd", "rd"][(day % 10 > 3 || Math.floor(day % 100 / 10) === 1) ? 0 : day % 10];
                        dayDisplay = `${day}${suffix}`;
                      } else if (user.Birthday) {
                        dayDisplay = user.Birthday;
                      }

                      const userImg = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${user.UserEmail}`;

                      return (
                        <div key={idx} className={styles.userCard}>
                          <div className={styles.avatarWrapper}>
                            <img 
                              src={userImg} 
                              onError={(e) => { 
                                (e.target as HTMLImageElement).src = 'https://office.net'; 
                              }}
                              className={styles.avatar} 
                              alt={user.Name} 
                            />
                            <div className={styles.balloonMini}>🎈</div>
                          </div>
                          <div className={styles.userInfo}>
                            <div className={styles.userName}>{user.Name}</div>
                            <div className={styles.userDate}>Birthday: <span>{dayDisplay}</span></div>
                          </div>
                          <a 
                            href={`mailto:${user.UserEmail}?subject=Happy Birthday! 🎉`} 
                            className={styles.wishButton}
                            title={`Send a birthday wish to ${user.Name}`}
                          >
                            Wish 🎉
                          </a>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}


