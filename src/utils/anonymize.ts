import type { StandardField } from './excelUtils';

const FIRST_NAMES = [
  'James', 'Emma', 'Oliver', 'Sophia', 'William', 'Ava', 'Benjamin', 'Isabella',
  'Lucas', 'Mia', 'Henry', 'Charlotte', 'Alexander', 'Amelia', 'Mason', 'Harper',
  'Ethan', 'Evelyn', 'Daniel', 'Abigail', 'Michael', 'Emily', 'Logan', 'Elizabeth',
  'Jackson', 'Mila', 'Sebastian', 'Ella', 'Jack', 'Scarlett', 'Aiden', 'Grace',
  'Owen', 'Chloe', 'Samuel', 'Victoria', 'Matthew', 'Riley', 'Joseph', 'Aria',
  'Liam', 'Lily', 'Noah', 'Layla', 'Elijah', 'Zoe', 'Jayden', 'Natalie',
  'Gabriel', 'Madison', 'Carter', 'Hannah', 'Julian', 'Addison', 'Wyatt', 'Aubrey',
  'Luke', 'Ellie', 'Isaac', 'Stella', 'Dylan', 'Violet', 'Anthony', 'Penelope',
  'Leo', 'Claire', 'Lincoln', 'Aurora', 'Jaxon', 'Nora', 'Asher', 'Skylar',
  'Christopher', 'Sofia', 'Joshua', 'Eleanor', 'Andrew', 'Paisley', 'Caleb', 'Savannah',
  'Ryan', 'Anna', 'Nathan', 'Hazel', 'Aaron', 'Isla', 'Christian', 'Willow',
  'Landon', 'Leah', 'Hunter', 'Lillian', 'Connor', 'Lucy', 'Eli', 'Alice',
  'David', 'Bella', 'Charlie', 'Brooklyn', 'Jonathan', 'Alexa', 'Colton', 'Naomi',
  'Evan', 'Caroline', 'Hudson', 'Elena', 'Dominic', 'Maya', 'Tucker', 'Julia',
  'Xavier', 'Ariana', 'Levi', 'Aaliyah', 'Adrian', 'Madelyn', 'Gavin', 'Eva',
  'Nolan', 'Quinn', 'Camden', 'Piper', 'Tyler', 'Serenity', 'Kayden', 'Valentina',
  'Robert', 'Lydia', 'Brayden', 'Eliana', 'Jordan', 'Marcus', 'Isabel', 'Finn',
  'Zara', 'Oscar', 'Freya', 'Theo', 'Ivy', 'Felix', 'Ruby',
];

const LAST_NAMES = [
  'Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis',
  'Rodriguez', 'Martinez', 'Hernandez', 'Lopez', 'Gonzalez', 'Wilson', 'Anderson',
  'Thomas', 'Taylor', 'Moore', 'Jackson', 'Martin', 'Lee', 'Perez', 'Thompson',
  'White', 'Harris', 'Sanchez', 'Clark', 'Ramirez', 'Lewis', 'Robinson', 'Walker',
  'Young', 'Allen', 'King', 'Wright', 'Scott', 'Torres', 'Nguyen', 'Hill', 'Flores',
  'Green', 'Adams', 'Nelson', 'Baker', 'Hall', 'Rivera', 'Campbell', 'Mitchell',
  'Carter', 'Roberts', 'Turner', 'Phillips', 'Evans', 'Collins', 'Edwards', 'Stewart',
  'Morris', 'Murphy', 'Cook', 'Rogers', 'Gutierrez', 'Ortiz', 'Morgan', 'Cooper',
  'Peterson', 'Bailey', 'Reed', 'Kelly', 'Howard', 'Ramos', 'Kim', 'Cox', 'Ward',
  'Richardson', 'Watson', 'Brooks', 'Chavez', 'Wood', 'James', 'Bennett', 'Gray',
  'Mendoza', 'Ruiz', 'Hughes', 'Price', 'Alvarez', 'Castillo', 'Sanders', 'Patel',
  'Myers', 'Long', 'Ross', 'Foster', 'Jimenez', 'Powell', 'Jenkins', 'Perry',
  'Russell', 'Sullivan', 'Bell', 'Coleman', 'Butler', 'Henderson', 'Barnes', 'Fisher',
  'Vasquez', 'Simmons', 'Romero', 'Jordan', 'Patterson', 'Alexander', 'Hamilton',
  'Graham', 'Reynolds', 'Griffin', 'Wallace', 'Moreno', 'West', 'Cole', 'Hayes',
  'Bryant', 'Herrera', 'Gibson', 'Ford', 'Ellis', 'Harrison', 'Stone', 'Murray',
  'Marshall', 'Owens', 'McDonald', 'Kennedy', 'Wells', 'Dixon', 'Robertson', 'Black',
];

const buildFakeName = (index: number): string => {
  const first = FIRST_NAMES[index % FIRST_NAMES.length];
  const last = LAST_NAMES[Math.floor(index / FIRST_NAMES.length) % LAST_NAMES.length];
  return `${first} ${last}`;
};

const normalizeKey = (value: string): string =>
  value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '');

export const anonymizeStudents = (rows: StandardField[]): StandardField[] => {
  const nameMap = new Map<string, string>();
  let nextIndex = 0;

  return rows.map((row) => {
    if (!row.navn) {
      return row;
    }

    const key = normalizeKey(row.navn);
    let fakeName = nameMap.get(key);
    if (!fakeName) {
      fakeName = buildFakeName(nextIndex);
      nextIndex += 1;
      nameMap.set(key, fakeName);
    }

    return { ...row, navn: fakeName };
  });
};
