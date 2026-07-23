import * as React from 'react';
import { useState, useEffect } from 'react';
import { IBirthdayProps } from './IBirthdayProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
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
    fetchBirthdays();
  }, []);

    const fetchBirthdays = async (): Promise<void> => {
    try {
      setApiError(null);
      
      // Hardcoded subsite path to guarantee it hits the correct list location
      const subsiteUrl = "https://tns05.sharepoint.com/sites/ProjectManagement";
      const listUrl = `${subsiteUrl}/_api/web/lists/GetByTitle('Birthdays')/items?$select=Birthday,Name/Title,Name/EMail&$expand=Name`;
      
      console.log("Requesting data from URL:", listUrl);
      let response = await props.context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
      
      if (!response.ok) {
        throw new Error(`SharePoint REST API Error (Status ${response.status}). Please make sure a list named 'Birthdays' exists at ${subsiteUrl}`);
      }

      const data = await response.json();
      console.log("Successfully fetched raw SharePoint payload:", data.value);

      if (data.value && data.value.length > 0) {
        const processedData: IBirthdayUser[] = data.value
          .filter((item: any) => item.Name) 
          .map((item: any) => ({
            Name: item.Name.Title || "Unknown User",
            Birthday: item.Birthday ? item.Birthday.toString() : "", 
            UserEmail: item.Name.EMail || ""
          }));

        groupUsersByMonth(processedData);
      } else {
        console.warn("The SharePoint 'Birthdays' list turned up empty.");
      }
      
      setLoading(false);
    } catch (error) {
      console.error('Fatal data fetching process crash:', error);
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

      // Plan A: Check for native ISO calendar system structures
      const parsedDate = new Date(user.Birthday);
      if (!isNaN(parsedDate.getTime())) {
        monthIndex = parsedDate.getMonth();
      } else {
        // Plan B: Text string keyword fallback if the row holds manually typed text entries
        const lowerBirthday = user.Birthday.toLowerCase();
        monthIndex = MONTHS.findIndex(m => lowerBirthday.includes(m.toLowerCase()));
      }

      if (monthIndex >= 0 && monthIndex < 12) {
        const monthName = MONTHS[monthIndex];
        initialGroups[monthName].push(user);
      } else {
        console.warn(`Could not parse calendar timeline index for user ${user.Name} from value: "${user.Birthday}"`);
      }
    });

    setGroupedBirthdays(initialGroups);
  };

  const toggleMonth = (month: string): void => {
    setExpandedMonths(prev => ({
      ...prev,
      [month]: !prev[month]
    }));
  };

  if (loading) {
    return <div className={styles.loading}>Loading site birthday registry database entries...</div>;
  }

  if (apiError) {
    return (
      <div className={styles.birthdayWebPart} style={{ border: '2px dashed #a80000', padding: '15px' }}>
        <h3 style={{ color: '#a80000', margin: '0 0 10px 0' }}>Data Connection Issue</h3>
        <p style={{ margin: 0, color: '#323130' }}>{apiError}</p>
        <p style={{ fontSize: '12px', marginTop: '10px', color: '#605e5c' }}>
          Target Web URL Context: <code>{props.context.pageContext.web.absoluteUrl}</code>
        </p>
      </div>
    );
  }

  return (
    <div className={styles.birthdayWebPart}>
      <h2 className={styles.title}>Birthdays Per Month</h2>
      
      {MONTHS.map(month => {
        const usersInMonth = groupedBirthdays[month] || [];
        const isExpanded = !!expandedMonths[month];

        return (
          <div key={month} className={styles.accordionSection}>
            <button 
              className={`${styles.accordionHeader} ${isExpanded ? styles.active : ''}`}
              onClick={() => toggleMonth(month)}
            >
              <span>{month}</span>
              <span className={styles.icon}>{isExpanded ? '▼' : '►'}</span>
            </button>
            
            {isExpanded && (
              <div className={styles.accordionContent}>
                {usersInMonth.length === 0 ? (
                  <p className={styles.noBirthdays}>No birthdays this month.</p>
                ) : (
                  <div className={styles.userGrid}>
                    {usersInMonth.map((user, idx) => {
                      let dayDisplay = "";
                      const dateObj = new Date(user.Birthday);
                      if (!isNaN(dateObj.getTime())) {
                        dayDisplay = `Day: ${dateObj.getDate()}`;
                      } else {
                        dayDisplay = user.Birthday; 
                      }

                      const userImg = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${user.UserEmail}`;

                      return (
                        <div key={idx} className={styles.userCard}>
                          <img 
                            src={userImg} 
                            onError={(e) => { 
                              (e.target as HTMLImageElement).src = 'https://office.net'; 
                            }}
                            className={styles.avatar} 
                            alt={user.Name} 
                          />
                          <div className={styles.userInfo}>
                            <div className={styles.userName}>{user.Name}</div>
                            <div className={styles.userDate}>{dayDisplay}</div>
                          </div>
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

