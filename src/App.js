import React, {
  useState,
  useEffect,
  createContext,
  useContext,
  useMemo,
} from 'react';
import './App.css'; // Keep this import

const CourseContext = createContext();

const CourseProvider = ({ children }) => {
  const [courses, setCourses] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchProgress, setFetchProgress] = useState(0); // Track pages fetched
  const [error, setError] = useState(null);
  const totalPages = 575; // Store total number of pages for progress calculation

  useEffect(() => {
    const fetchAllCourses = async () => {
      setLoading(true);
      setError(null);
      setFetchProgress(0);

      try {
        const allCoursesData = [];
        for (let page = 1; page <= totalPages; page++) {
          try {  //Try/Catch for *each* page fetch
            const apiUrl = `https://universitaly-backend.cineca.it/api/offerta-formativa/cerca-corsi?searchType=u&page=${page}&tipoLaurea&tipoClasse=0&durata&lingua&tipoAccesso&modalitaErogazione&searchText&area&order=RND&provincia&provinciaSigla`;
            const response = await fetch(apiUrl);

            if (!response.ok) {
              throw new Error(
                `Failed to fetch data for page ${page}: ${response.status}`,
              );
            }

            const data = await response.json();
            allCoursesData.push(...data.corsi);
            setFetchProgress((prev) => prev + 1); // Increment progress
          } catch (pageFetchError) {
            // Log the error, but continue fetching other pages
            console.error(`Error fetching page ${page}:`, pageFetchError);
            setError(`Error fetching page ${page}: ${pageFetchError.message}`);  //Set Error Message
          }
        }

        setCourses(allCoursesData);
      } catch (err) {
        setError(err.message);
        console.error('Error fetching courses:', err);
      } finally {
        setLoading(false);
      }
    };

    fetchAllCourses();
  }, []);

  const value = useMemo(
    () => ({
      courses,
      loading,
      fetchProgress,
      totalPages,
      error,
    }),
    [courses, loading, fetchProgress, totalPages, error],
  );

  return (
    <CourseContext.Provider value={value}>{children}</CourseContext.Provider>
  );
};

const useCourses = () => {
  const context = useContext(CourseContext);
  if (!context) {
    throw new Error('useCourses must be used within a CourseProvider');
  }
  return context;
};

const CourseFilter = () => {
  const { courses, loading, fetchProgress, totalPages, error } = useCourses();
  const [searchTerm, setSearchTerm] = useState('');
  const [englishOnly, setEnglishOnly] = useState(false);
  const [degreeType, setDegreeType] = useState(''); // '' means no filter

  const filteredCourses = useMemo(() => {
    let filtered = courses;

    if (searchTerm) {
      filtered = filtered.filter((course) =>
        course.nomeCorsoEn.toLowerCase().includes(searchTerm.toLowerCase()),
      );
    }

    if (englishOnly) {
      filtered = filtered.filter((course) => course.lingua === 'EN');
    }

    if (degreeType) {
      filtered = filtered.filter(
        (course) => course.tipoLaurea.descrizioneEn === degreeType,
      );
    }

    return filtered;
  }, [courses, searchTerm, englishOnly, degreeType]);

  const handleSearchChange = (event) => {
    setSearchTerm(event.target.value);
  };

  const handleEnglishOnlyChange = () => {
    setEnglishOnly(!englishOnly);
  };

  const handleDegreeTypeChange = (event) => {
    setDegreeType(event.target.value);
  };

  if (loading) {
    return (
      <div>
        <p>
          Loading courses... Fetched {fetchProgress} / {totalPages} pages.
        </p>
        {error && <p style={{ color: 'red' }}>{error}</p>}  {/* Display error during loading */}
      </div>
    );
  }

  if (error && courses.length === 0) {
      //Only display the error message if no courses have been loaded.
      return <div>Error: {error}</div>;
  }

  return (
    <div className="container">
      <h1>University Course Search</h1>
      <div className="filters">
        <input
          type="text"
          placeholder="Search by course name..."
          value={searchTerm}
          onChange={handleSearchChange}
        />
        <label>
          <input
            type="checkbox"
            checked={englishOnly}
            onChange={handleEnglishOnlyChange}
          />
          English Only
        </label>
        <select value={degreeType} onChange={handleDegreeTypeChange}>
          <option value="">All Degree Types</option>
          <option value="EN Triennale">Bachelor's</option>
          <option value="EN Magistrale">Master's</option>
        </select>
      </div>

      <h2>Results:</h2>
      <table>
        <thead>
          <tr>
            <th>Course Name</th>
            <th>Degree Type</th>
            <th>Language</th>
          </tr>
        </thead>
        <tbody>
          {filteredCourses.map((course) => (
            <tr key={course.id} className="course-item">
              <td>{course.nomeCorsoEn}</td>
              <td>{course.tipoLaurea.descrizioneEn}</td>
              <td>{course.lingua}</td>
            </tr>
          ))}
        </tbody>
      </table>

    </div>
  );
};

const App = () => {
  return (
    <CourseProvider>
      <CourseFilter />
    </CourseProvider>
  );
};

export default App;