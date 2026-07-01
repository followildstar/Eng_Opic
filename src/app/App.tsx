import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Home, LayoutList, RotateCcw, BarChart2, Search, Star,
  ChevronDown, ChevronUp, Mic, Square, ArrowLeft,
  Shuffle, SkipBack, SkipForward, Eye, EyeOff, Check,
  Flame, Clock, Target, BookOpen, ChevronRight, Plus,
  Trash2, X, Download
} from "lucide-react";

// ─── Types ───────────────────────────────────────────────────────────────────
type PracticeLog = Record<string, number>; // "2026-06-30" → 그날 연습 분(分)
type Screen = "home" | "questions" | "detail" | "practice" | "shadowing" | "review" | "progress" | "add" | "edit";
type Status = "New" | "Practicing" | "Confident";
type TabId = "home" | "questions" | "review" | "progress";
type Category = "Custom" | "Intro" | "Hobby" | "Routine" | "Mindset" | "Experience" | "Place" | "Study";

interface Question {
  id: number;
  title: string;
  category: Category;
  set: string;
  status: Status;
  favorite: boolean;
  flowSteps: string[];
  keyWords: string[];
  fullAnswer: string;
  fullAnswerSentences: string[];
  similarQuestions: string[];
  lastPracticed?: string;
  practiceCount: number;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function todayKey(): string {
  return new Date().toISOString().slice(0, 10); // "2026-06-30"
}

function dateKeyFrom(daysAgo: number): string {
  const d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return d.toISOString().slice(0, 10);
}

function shortDayLabel(daysAgo: number): string {
  const d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return ["S", "M", "T", "W", "T", "F", "S"][d.getDay()];
}

function shortDateLabel(): string {
  return new Date().toLocaleDateString("en-US", { month: "short", day: "numeric" });
}
function splitSentences(text: string): string[] {
  const matches = text.match(/[^.!?]+[.!?]+[\s]*/g);
  return matches ? matches.map(s => s.trim()).filter(Boolean) : [text.trim()];
}

function FmtTime(secs: number) {
  const m = Math.floor(secs / 60).toString().padStart(2, "0");
  const s = (secs % 60).toString().padStart(2, "0");
  return `${m}:${s}`;
}

// ─── Data ────────────────────────────────────────────────────────────────────

const INITIAL_QUESTIONS: Question[] = [
  {
    id: 1, title: "Please introduce yourself.", category: "Intro", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Name and job", "Creative hobbies", "Healthy habits and growth", "English goal", "Current life attitude"],
    keyWords: ["web designer", "drawing", "piano", "personal growth", "express my thoughts"],
    fullAnswer: "Hi, my name is Amy. I work as a web designer and publisher in Korea. I enjoy creative hobbies like drawing and playing the piano. Recently, I've been trying to build healthier habits and focus more on personal growth. I also study English because I want to express my thoughts more naturally. These days, I'm trying to enjoy simple things in life and take care of myself more.",
    fullAnswerSentences: ["Hi, my name is Amy.", "I work as a web designer and publisher in Korea.", "I enjoy creative hobbies like drawing and playing the piano.", "Recently, I've been trying to build healthier habits and focus more on personal growth.", "I also study English because I want to express my thoughts more naturally.", "These days, I'm trying to enjoy simple things in life and take care of myself more."],
    similarQuestions: ["Tell me about yourself.", "What kind of person are you?", "Tell me about your daily life.", "What are your hobbies?", "What are you interested in these days?"],
  },
  {
    id: 2, title: "Tell me about your hobbies.", category: "Hobby", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Drawing and piano", "Drawing expresses myself", "Stopped because of perfectionism", "Restarted with process mindset", "Piano helps me relax"],
    keyWords: ["drawing", "piano", "perfection", "process", "Turkish March"],
    fullAnswer: "I enjoy drawing and playing the piano. Drawing is something I really love because it helps me express myself. Actually, I stopped drawing for a while because I wanted everything to be perfect. But recently, I started again and realized enjoying the process is more important. I also enjoy piano practice. Lately, I've been practicing the Turkish March. Playing piano helps me relax and makes me feel proud of myself.",
    fullAnswerSentences: ["I enjoy drawing and playing the piano.", "Drawing is something I really love because it helps me express myself.", "Actually, I stopped drawing for a while because I wanted everything to be perfect.", "But recently, I started again and realized enjoying the process is more important.", "I also enjoy piano practice.", "Lately, I've been practicing the Turkish March.", "Playing piano helps me relax and makes me feel proud of myself."],
    similarQuestions: ["Tell me about your favorite hobby.", "What do you usually do in your free time?", "Describe something you enjoy doing.", "Tell me about a hobby you recently started or restarted.", "What helps you relax?"],
  },
  {
    id: 3, title: "What do you usually do in your free time?", category: "Hobby", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Quiet activities", "Drawing, piano, workout, yoga", "Writing and organizing thoughts", "Recharge", "Small routines feel meaningful"],
    keyWords: ["quiet activities", "recharge", "meaningful routines"],
    fullAnswer: "In my free time, I usually do quiet activities. I enjoy drawing, practicing piano, working out, or going to yoga classes. Sometimes I also spend time organizing my thoughts or writing. I like peaceful activities because they help me recharge. I think small routines make my daily life feel meaningful.",
    fullAnswerSentences: ["In my free time, I usually do quiet activities.", "I enjoy drawing, practicing piano, working out, or going to yoga classes.", "Sometimes I also spend time organizing my thoughts or writing.", "I like peaceful activities because they help me recharge.", "I think small routines make my daily life feel meaningful."],
    similarQuestions: ["What do you usually do on weekends?", "How do you spend your free time?", "What do you do when you want to relax?", "Tell me about your hobbies.", "What do you do after work?"],
  },
  {
    id: 4, title: "Tell me about your exercise routine.", category: "Routine", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Regular exercise", "Yoga and simple workouts", "Body and mind", "Proud and energetic", "Consistency over perfection"],
    keyWords: ["regularly", "yoga", "mental health", "consistent"],
    fullAnswer: "I try to exercise regularly. These days, I mostly do yoga and simple workouts. I learned that exercise is really important for both physical and mental health. After working out, I feel proud of myself and more energetic. It also helps me focus better during the day. Recently, I've been trying to stay consistent instead of being perfect.",
    fullAnswerSentences: ["I try to exercise regularly.", "These days, I mostly do yoga and simple workouts.", "I learned that exercise is really important for both physical and mental health.", "After working out, I feel proud of myself and more energetic.", "It also helps me focus better during the day.", "Recently, I've been trying to stay consistent instead of being perfect."],
    similarQuestions: ["What do you do to stay healthy?", "Tell me about your workout routine.", "How often do you exercise?", "Describe a healthy habit you have.", "What helps you manage stress?"],
  },
  {
    id: 5, title: "When do you usually work out?", category: "Routine", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Prefer morning", "Difficult at first", "Morning is consistent", "Evening changes easily", "Productive and confident"],
    keyWords: ["morning workout", "consistent", "evening variables", "confidence"],
    fullAnswer: "I usually prefer morning workouts. At first, it was difficult because I felt tired. But later, I realized mornings are more consistent for me. In the evening, plans can easily change or I can feel too tired. Morning exercise helps me feel more productive and confident throughout the day.",
    fullAnswerSentences: ["I usually prefer morning workouts.", "At first, it was difficult because I felt tired.", "But later, I realized mornings are more consistent for me.", "In the evening, plans can easily change or I can feel too tired.", "Morning exercise helps me feel more productive and confident throughout the day."],
    similarQuestions: ["What is your morning routine?", "Why do you like mornings?", "When are you most productive?", "Tell me about a recent change in your routine.", "Why do you prefer morning exercise?"],
  },
  {
    id: 6, title: "Tell me about a recent change in your life.", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Mindset changed", "Perfection focus before", "Stopped what I loved", "Now consistency and process", "Happier and less stressed"],
    keyWords: ["mindset", "perfection", "consistency", "process"],
    fullAnswer: "Recently, my mindset changed a lot. Before, I focused too much on perfection. Because of that, I sometimes stopped doing things I actually loved. But now, I focus more on consistency and enjoying the process. I think this mindset made me happier and less stressed.",
    fullAnswerSentences: ["Recently, my mindset changed a lot.", "Before, I focused too much on perfection.", "Because of that, I sometimes stopped doing things I actually loved.", "But now, I focus more on consistency and enjoying the process.", "I think this mindset made me happier and less stressed."],
    similarQuestions: ["How have you changed recently?", "Tell me about a change you made.", "What changed in your mindset?", "Describe a recent personal growth experience.", "Tell me about something important you realized."],
  },
  {
    id: 7, title: "Tell me about something you recently learned.", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Learned about myself", "No need to be perfect", "Drawing and others' opinions", "Draw what I like", "Creative and confident"],
    keyWords: ["not perfect", "what I like", "creative confidence"],
    fullAnswer: "Recently, I learned something important about myself. I realized I don't need to be perfect to enjoy something. For example, with drawing, I used to worry too much about what other people thought. But now, I just draw what I truly like. That change made me feel more creative and confident.",
    fullAnswerSentences: ["Recently, I learned something important about myself.", "I realized I don't need to be perfect to enjoy something.", "For example, with drawing, I used to worry too much about what other people thought.", "But now, I just draw what I truly like.", "That change made me feel more creative and confident."],
    similarQuestions: ["Tell me about an important lesson you learned.", "What did you learn about yourself recently?", "Describe something meaningful you realized.", "Tell me about a personal discovery.", "What helped you grow recently?"],
  },
  {
    id: 8, title: "What do you do when you feel stressed?", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Take care of myself", "Workout, piano, walk", "Write thoughts", "Do not avoid emotions", "Process in healthy ways"],
    keyWords: ["stress", "take care", "process feelings", "calmer"],
    fullAnswer: "When I feel stressed, I try to take care of myself. Usually, I work out, play piano, or go for a walk. Sometimes I write down my thoughts too. I learned that avoiding emotions doesn't really help. Instead, I try to process my feelings in healthy ways. After that, I usually feel calmer.",
    fullAnswerSentences: ["When I feel stressed, I try to take care of myself.", "Usually, I work out, play piano, or go for a walk.", "Sometimes I write down my thoughts too.", "I learned that avoiding emotions doesn't really help.", "Instead, I try to process my feelings in healthy ways.", "After that, I usually feel calmer."],
    similarQuestions: ["How do you manage stress?", "What helps you relax?", "What do you do when you feel overwhelmed?", "Tell me about your stress relief habits.", "How do you take care of yourself?"],
  },
  {
    id: 9, title: "Describe your daily routine.", category: "Routine", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Simple weekdays", "Wake up and work", "After work habits", "Workout, piano, English, drawing", "Routine gives stability"],
    keyWords: ["weekday", "after work", "healthy habits", "stable"],
    fullAnswer: "My weekdays are pretty simple. I usually wake up, get ready, and go to work. After work, I try to focus on healthy habits. I work out, practice piano, study English, or draw. Before sleeping, I sometimes listen to English audio or review what I studied. I like routines because they make me feel stable.",
    fullAnswerSentences: ["My weekdays are pretty simple.", "I usually wake up, get ready, and go to work.", "After work, I try to focus on healthy habits.", "I work out, practice piano, study English, or draw.", "Before sleeping, I sometimes listen to English audio or review what I studied.", "I like routines because they make me feel stable."],
    similarQuestions: ["Tell me about your weekday routine.", "What does a typical day look like for you?", "What do you usually do after work?", "Describe your lifestyle.", "What habits are important to you?"],
  },
  {
    id: 10, title: "What do you usually do on weekends?", category: "Routine", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Relax and recharge", "Workout, piano, drawing", "Library or quiet cafe", "Slow mornings", "Time with my cat"],
    keyWords: ["weekend", "recharge", "slow morning", "cat"],
    fullAnswer: "On weekends, I try to relax and recharge. I usually work out more, practice piano, and spend time drawing. Sometimes I go to a library or a quiet café. I enjoy slow mornings on weekends because they make me feel peaceful. I also like spending quiet time with my cat at home.",
    fullAnswerSentences: ["On weekends, I try to relax and recharge.", "I usually work out more, practice piano, and spend time drawing.", "Sometimes I go to a library or a quiet café.", "I enjoy slow mornings on weekends because they make me feel peaceful.", "I also like spending quiet time with my cat at home."],
    similarQuestions: ["How do you spend your weekends?", "Tell me about your free time.", "Describe a relaxing day.", "What do you usually do when you are off work?", "What do you enjoy doing on weekends?"],
  },
  {
    id: 11, title: "Tell me about a memorable experience recently.", category: "Experience", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Yoga experience", "Saw someone's hand", "Looked at my hands", "Felt grateful for health", "Appreciate what I have"],
    keyWords: ["yoga", "hands", "health", "gratitude"],
    fullAnswer: "Recently, I had a meaningful experience during yoga. I saw someone with a disability and suddenly looked at my own hands. At that moment, I felt grateful for being healthy. It made me realize I should appreciate what I already have. Since then, I've been trying to live with more gratitude.",
    fullAnswerSentences: ["Recently, I had a meaningful experience during yoga.", "I saw someone with a disability and suddenly looked at my own hands.", "At that moment, I felt grateful for being healthy.", "It made me realize I should appreciate what I already have.", "Since then, I've been trying to live with more gratitude."],
    similarQuestions: ["Describe a meaningful experience.", "Tell me about a moment that changed your perspective.", "Tell me about a recent realization.", "Describe a memorable moment.", "What made you feel grateful recently?"],
  },
  {
    id: 12, title: "Tell me about a hobby you stopped and restarted.", category: "Hobby", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Drawing restarted", "Stopped due to perfection", "Cared about opinions", "Started again", "Draw for enjoyment"],
    keyWords: ["drawing", "stopped", "restarted", "enjoy", "not impress"],
    fullAnswer: "One hobby I restarted is drawing. I stopped for a while because I worried too much about doing it perfectly. I also cared too much about what other people thought. But recently, I started again. Now, I draw what I truly enjoy instead of trying to impress people. It feels much more fun and natural now.",
    fullAnswerSentences: ["One hobby I restarted is drawing.", "I stopped for a while because I worried too much about doing it perfectly.", "I also cared too much about what other people thought.", "But recently, I started again.", "Now, I draw what I truly enjoy instead of trying to impress people.", "It feels much more fun and natural now."],
    similarQuestions: ["Tell me about something you started again.", "Describe a hobby that is important to you.", "Why did you stop your hobby?", "Tell me about something you changed recently.", "Describe something you enjoy these days."],
  },
  {
    id: 13, title: "What makes you happy these days?", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Simple things", "Workout, piano, drawing, English", "Life I dreamed of", "Gratitude", "Enjoy present"],
    keyWords: ["simple happiness", "dreamed life", "present"],
    fullAnswer: "Simple things make me happy these days. Working out, playing piano, drawing, and studying English. Recently, I realized I'm already living the life I once dreamed about. That made me feel much more grateful. Now, I try to enjoy the present instead of always chasing the next goal.",
    fullAnswerSentences: ["Simple things make me happy these days.", "Working out, playing piano, drawing, and studying English.", "Recently, I realized I'm already living the life I once dreamed about.", "That made me feel much more grateful.", "Now, I try to enjoy the present instead of always chasing the next goal."],
    similarQuestions: ["What do you enjoy these days?", "What makes your life meaningful?", "What are you grateful for these days?", "Tell me about a recent happy moment.", "What gives you energy these days?"],
  },
  {
    id: 14, title: "What are your future goals?", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Healthier and creative", "Improve English and creative skills", "Drawing and piano", "Meaningful balanced life", "Enjoy process and grow"],
    keyWords: ["future goal", "healthier", "creative", "balanced"],
    fullAnswer: "My goal is to become healthier and more creative. I want to keep improving my English and creative skills. I also want to continue drawing and playing piano. Most importantly, I want to live a meaningful and balanced life. I want to become someone who enjoys the process and keeps growing.",
    fullAnswerSentences: ["My goal is to become healthier and more creative.", "I want to keep improving my English and creative skills.", "I also want to continue drawing and playing piano.", "Most importantly, I want to live a meaningful and balanced life.", "I want to become someone who enjoys the process and keeps growing."],
    similarQuestions: ["Tell me about your future plans.", "What do you want to achieve?", "What are your personal goals?", "What kind of life do you want in the future?", "What is important to you?"],
  },
  {
    id: 15, title: "How have you changed over time?", category: "Mindset", set: "Set 01",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Changed a lot", "Perfection and opinions before", "Stress", "Focus on myself", "Consistency over perfection"],
    keyWords: ["changed", "perfection", "opinions", "calmer", "confident"],
    fullAnswer: "I think I changed a lot over time. Before, I cared too much about perfection and other people's opinions. Because of that, I often felt stressed. But now, I try to focus more on myself and what truly matters. I learned that consistency is more important than perfection. I think I'm becoming calmer and more confident.",
    fullAnswerSentences: ["I think I changed a lot over time.", "Before, I cared too much about perfection and other people's opinions.", "Because of that, I often felt stressed.", "But now, I try to focus more on myself and what truly matters.", "I learned that consistency is more important than perfection.", "I think I'm becoming calmer and more confident."],
    similarQuestions: ["Describe how you changed.", "What kind of person were you before?", "Tell me about personal growth.", "How is your life different now?", "Tell me about an important change in your life."],
  },
  {
    id: 16, title: "Tell me about a difficult time in your life and how you handled it.", category: "Experience", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Perfectionism was difficult", "Stopped what I loved", "Drawing example", "Fear of disappointment", "Changed mindset", "Process and starting"],
    keyWords: ["difficult time", "perfectionism", "drawing", "process"],
    fullAnswer: "One difficult thing for me was dealing with perfectionism. I used to stop doing things I loved because I wanted everything to be perfect. For example, I stopped drawing for a long time. I was afraid of disappointing myself or not being good enough. But recently, I changed my mindset. Now, I focus more on enjoying the process and just starting. That change helped me feel happier and less stressed.",
    fullAnswerSentences: ["One difficult thing for me was dealing with perfectionism.", "I used to stop doing things I loved because I wanted everything to be perfect.", "For example, I stopped drawing for a long time.", "I was afraid of disappointing myself or not being good enough.", "But recently, I changed my mindset.", "Now, I focus more on enjoying the process and just starting.", "That change helped me feel happier and less stressed."],
    similarQuestions: ["Tell me about a challenge you faced.", "Describe a difficult experience.", "Tell me about something stressful in your life.", "How did you overcome a problem?", "Tell me about a time you changed yourself."],
  },
  {
    id: 17, title: "Tell me about something you recently realized.", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Realized happiness", "Compared myself before", "Felt behind", "Workout changed thought", "Dreamed life now", "Grateful"],
    keyWords: ["realization", "comparison", "workout", "dreamed life"],
    fullAnswer: "Recently, I realized something important about happiness. I used to compare myself to other people a lot. Sometimes I felt behind in life. But one day while working out, I suddenly realized something. I'm already living the life I once dreamed about. I have meaningful work, hobbies I enjoy, and healthier habits. That realization made me feel much more grateful.",
    fullAnswerSentences: ["Recently, I realized something important about happiness.", "I used to compare myself to other people a lot.", "Sometimes I felt behind in life.", "But one day while working out, I suddenly realized something.", "I'm already living the life I once dreamed about.", "I have meaningful work, hobbies I enjoy, and healthier habits.", "That realization made me feel much more grateful."],
    similarQuestions: ["Tell me about an important realization.", "What made you feel grateful recently?", "Describe something meaningful you learned.", "Tell me about a recent change in perspective.", "What changed your mindset?"],
  },
  {
    id: 18, title: "Describe a memorable place you often go to.", category: "Place", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Yoga studio", "Relax and clear mind", "Peaceful atmosphere", "Feel calmer after", "New life ideas", "Mental reset"],
    keyWords: ["yoga studio", "peaceful", "reset mentally"],
    fullAnswer: "One place I often go to is my yoga studio. I like it because it helps me relax and clear my mind. Usually, the atmosphere is very peaceful. After yoga, I feel calmer and healthier. Sometimes, I even get new ideas about my life during class. That place became important to me because it helps me reset mentally.",
    fullAnswerSentences: ["One place I often go to is my yoga studio.", "I like it because it helps me relax and clear my mind.", "Usually, the atmosphere is very peaceful.", "After yoga, I feel calmer and healthier.", "Sometimes, I even get new ideas about my life during class.", "That place became important to me because it helps me reset mentally."],
    similarQuestions: ["Tell me about a place you often visit.", "Describe a place where you relax.", "Tell me about your yoga studio.", "What helps you manage stress?", "Describe a meaningful place."],
  },
  {
    id: 19, title: "Tell me about a time you changed your opinion.", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Changed opinion about exercise", "Thought morning was not for me", "Tired before", "Tried again", "Morning better for consistency", "Consistency over timing"],
    keyWords: ["changed opinion", "morning workout", "consistency"],
    fullAnswer: "I changed my opinion about exercise. Before, I thought morning workouts were not for me. I felt tired and thought it didn't fit my lifestyle. But after trying again, I realized mornings are actually better for consistency. In the evening, I often get tired or busy. Morning workouts helped me feel more energetic and productive. So now, I think consistency matters more than perfect timing.",
    fullAnswerSentences: ["I changed my opinion about exercise.", "Before, I thought morning workouts were not for me.", "I felt tired and thought it didn't fit my lifestyle.", "But after trying again, I realized mornings are actually better for consistency.", "In the evening, I often get tired or busy.", "Morning workouts helped me feel more energetic and productive.", "So now, I think consistency matters more than perfect timing."],
    similarQuestions: ["Tell me about a time you changed your mind.", "Describe a change in your routine.", "What habit changed your life?", "Tell me about something that surprised you.", "What changed your perspective?"],
  },
  {
    id: 20, title: "Tell me about something you enjoy doing alone.", category: "Hobby", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Enjoy alone time", "Draw, piano, walk", "Organize thoughts", "Creative activities calm me", "Discover ideas", "Quiet time matters"],
    keyWords: ["alone time", "quiet", "organize thoughts", "creative"],
    fullAnswer: "I enjoy spending time alone a lot. Usually, I draw, practice piano, or go for a walk. I think being alone helps me organize my thoughts. Also, creative activities make me feel calm. Sometimes I even discover new ideas about myself. For me, quiet time is very important.",
    fullAnswerSentences: ["I enjoy spending time alone a lot.", "Usually, I draw, practice piano, or go for a walk.", "I think being alone helps me organize my thoughts.", "Also, creative activities make me feel calm.", "Sometimes I even discover new ideas about myself.", "For me, quiet time is very important."],
    similarQuestions: ["How do you spend time alone?", "What do you usually do in your free time?", "What helps you relax?", "Describe an activity you enjoy.", "What do you do when you need quiet time?"],
  },
  {
    id: 21, title: "Describe a recent problem and how you solved it.", category: "Experience", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Workout routine problem", "Morning exercise first", "Tired all day", "Needed adjustment", "Switched routine", "Learned flexibility"],
    keyWords: ["problem", "routine", "adjust", "flexibility"],
    fullAnswer: "Recently, I struggled with my workout routine. At first, I tried exercising in the morning. But after a while, I realized I felt tired all day. I thought something was wrong with me. Later, I realized I simply needed to adjust my routine. So I switched to evening workouts for a while. That experience taught me that flexibility is important.",
    fullAnswerSentences: ["Recently, I struggled with my workout routine.", "At first, I tried exercising in the morning.", "But after a while, I realized I felt tired all day.", "I thought something was wrong with me.", "Later, I realized I simply needed to adjust my routine.", "So I switched to evening workouts for a while.", "That experience taught me that flexibility is important."],
    similarQuestions: ["Tell me about a recent challenge.", "Describe a problem you had recently.", "Tell me about a time something did not go as planned.", "How did you solve a problem?", "Tell me about a change you made."],
  },
  {
    id: 22, title: "Tell me about someone important in your life.", category: "Experience", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Parents are important", "Appreciate them more", "Health and life realization", "Gratitude", "Express love more", "Family meaning"],
    keyWords: ["parents", "gratitude", "express love", "family"],
    fullAnswer: "My parents are very important to me. As I got older, I started appreciating them more. Recently, I realized how grateful I am for my health and life. That made me think more about my parents too. I want to express my love and gratitude more often. I think family becomes more meaningful as time passes.",
    fullAnswerSentences: ["My parents are very important to me.", "As I got older, I started appreciating them more.", "Recently, I realized how grateful I am for my health and life.", "That made me think more about my parents too.", "I want to express my love and gratitude more often.", "I think family becomes more meaningful as time passes."],
    similarQuestions: ["Who is important to you?", "Tell me about your family.", "Describe someone meaningful in your life.", "Who do you feel grateful for?", "Tell me about someone who influenced you."],
  },
  {
    id: 23, title: "Tell me about something you are trying to improve.", category: "Study", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Improving English", "Speak naturally and confidently", "Practice every day", "Listen, repeat, record", "Consistency over talent", "Feel improvement"],
    keyWords: ["English", "natural", "repeat", "record", "consistency"],
    fullAnswer: "These days, I'm trying to improve my English. I want to speak more naturally and confidently. So I practice every day. I listen to audio, repeat sentences, and record my voice. I think consistency is more important than talent. Little by little, I feel myself improving.",
    fullAnswerSentences: ["These days, I'm trying to improve my English.", "I want to speak more naturally and confidently.", "So I practice every day.", "I listen to audio, repeat sentences, and record my voice.", "I think consistency is more important than talent.", "Little by little, I feel myself improving."],
    similarQuestions: ["What are you learning these days?", "Tell me about a skill you want to improve.", "What do you do for self-improvement?", "Tell me about a recent goal.", "What is something challenging for you?"],
  },
  {
    id: 24, title: "Tell me about a healthy habit you have.", category: "Routine", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Regular exercise", "Body and mind", "Confidence and focus", "Manage stress", "Stay consistent", "Habits change life"],
    keyWords: ["healthy habit", "exercise", "stress", "consistency"],
    fullAnswer: "One healthy habit I'm trying to keep is regular exercise. I realized exercise affects both my body and mind. After working out, I feel more confident and focused. It also helps me manage stress better. Even when I don't feel motivated, I try to stay consistent. I think healthy habits slowly change your life.",
    fullAnswerSentences: ["One healthy habit I'm trying to keep is regular exercise.", "I realized exercise affects both my body and mind.", "After working out, I feel more confident and focused.", "It also helps me manage stress better.", "Even when I don't feel motivated, I try to stay consistent.", "I think healthy habits slowly change your life."],
    similarQuestions: ["What do you do to stay healthy?", "Tell me about your exercise routine.", "Describe a healthy habit.", "What helps you manage stress?", "Tell me about a good habit you have."],
  },
  {
    id: 25, title: "Tell me about something that changed your perspective.", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Drawing changed perspective", "Thought style was wrong", "Softer, prettier drawings", "Thought lack of skill", "Realized interpretation", "More confidence"],
    keyWords: ["drawing", "perspective", "interpretation", "confidence"],
    fullAnswer: "Drawing changed my perspective recently. I used to think my drawing style was wrong. Because my drawings looked softer or prettier than real life. I thought I lacked skill. But later, I realized it was actually my own artistic interpretation. That helped me feel more confident. Now, I enjoy drawing much more.",
    fullAnswerSentences: ["Drawing changed my perspective recently.", "I used to think my drawing style was wrong.", "Because my drawings looked softer or prettier than real life.", "I thought I lacked skill.", "But later, I realized it was actually my own artistic interpretation.", "That helped me feel more confident.", "Now, I enjoy drawing much more."],
    similarQuestions: ["Tell me about a realization you had.", "Describe something that changed the way you think.", "Tell me about a moment of growth.", "What helped you become more confident?", "Tell me about something meaningful you learned."],
  },
  {
    id: 26, title: "Describe a recent activity you enjoyed.", category: "Hobby", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Piano practice", "Rented grand piano room", "Compared myself at first", "Realized I work hard too", "Piano recharges me", "Meaningful experience"],
    keyWords: ["piano", "grand piano", "comparison", "recharge"],
    fullAnswer: "Recently, I enjoyed practicing piano. I rented a grand piano room and spent time practicing. At first, I felt a little awkward because I compared myself to better players. But while playing, I realized something. I've been working hard too. Piano helps me recharge after a long day. So the experience became very meaningful.",
    fullAnswerSentences: ["Recently, I enjoyed practicing piano.", "I rented a grand piano room and spent time practicing.", "At first, I felt a little awkward because I compared myself to better players.", "But while playing, I realized something.", "I've been working hard too.", "Piano helps me recharge after a long day.", "So the experience became very meaningful."],
    similarQuestions: ["Tell me about something fun you did recently.", "Describe a hobby you enjoy.", "Tell me about a meaningful experience.", "What do you enjoy doing in your free time?", "Describe a memorable activity."],
  },
  {
    id: 27, title: "Tell me about a personal goal you have.", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Healthier and balanced", "Body and mind", "Regular workout and routines", "Creative hobbies", "Growth means happier little by little"],
    keyWords: ["personal goal", "balanced", "body and mind", "growth"],
    fullAnswer: "One of my personal goals is becoming healthier and more balanced. I want to take care of both my body and mind. So I try to work out regularly and build healthy routines. I also want to keep improving my creative hobbies. For me, growth means becoming healthier and happier little by little.",
    fullAnswerSentences: ["One of my personal goals is becoming healthier and more balanced.", "I want to take care of both my body and mind.", "So I try to work out regularly and build healthy routines.", "I also want to keep improving my creative hobbies.", "For me, growth means becoming healthier and happier little by little."],
    similarQuestions: ["Tell me about your future goals.", "What kind of life do you want?", "What are you trying to improve?", "Describe something important to you.", "What motivates you these days?"],
  },
  {
    id: 28, title: "Describe something you learned from experience.", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Planning is not enough", "Too much thinking before", "Action is important", "Start before ready", "Reality feels real", "Changed approach"],
    keyWords: ["planning", "action", "start", "before ready"],
    fullAnswer: "I learned that planning alone is not enough. Before, I spent too much time thinking and planning. But recently, I realized action is more important. Even if I'm not fully ready, I try to start anyway. Once I actually do something, it feels more real. That mindset changed the way I approach goals.",
    fullAnswerSentences: ["I learned that planning alone is not enough.", "Before, I spent too much time thinking and planning.", "But recently, I realized action is more important.", "Even if I'm not fully ready, I try to start anyway.", "Once I actually do something, it feels more real.", "That mindset changed the way I approach goals."],
    similarQuestions: ["Tell me about an important lesson you learned.", "Describe something experience taught you.", "Tell me about a recent realization.", "What changed the way you think?", "Tell me about personal growth."],
  },
  {
    id: 29, title: "Tell me about a time you felt proud of yourself.", category: "Experience", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Proud after exercise consistency", "Difficult at first", "Positive changes", "More energy and confidence", "Trusted myself", "Proud feeling"],
    keyWords: ["proud", "exercise", "consistency", "self-trust"],
    fullAnswer: "Recently, I felt proud after staying consistent with exercise. At first, it was really difficult. But after a few weeks, I noticed positive changes. I had more energy and felt more confident. Most importantly, I started trusting myself more. That feeling made me proud.",
    fullAnswerSentences: ["Recently, I felt proud after staying consistent with exercise.", "At first, it was really difficult.", "But after a few weeks, I noticed positive changes.", "I had more energy and felt more confident.", "Most importantly, I started trusting myself more.", "That feeling made me proud."],
    similarQuestions: ["Describe something you achieved.", "Tell me about a proud moment.", "What made you feel confident recently?", "Tell me about a personal success.", "Describe a meaningful accomplishment."],
  },
  {
    id: 30, title: "What kind of person do you want to become?", category: "Mindset", set: "Set 02",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Calm, healthy, creative", "Express through art", "Speak thoughtfully", "Words have meaning", "Enjoy life and grow"],
    keyWords: ["calm", "healthy", "creative", "thoughtful", "grow"],
    fullAnswer: "I want to become someone calm, healthy, and creative. I want to express myself through drawing, piano, and writing. I also want to be someone who speaks thoughtfully. Instead of talking too much, I want my words to have meaning. Most importantly, I want to enjoy life while continuing to grow.",
    fullAnswerSentences: ["I want to become someone calm, healthy, and creative.", "I want to express myself through drawing, piano, and writing.", "I also want to be someone who speaks thoughtfully.", "Instead of talking too much, I want my words to have meaning.", "Most importantly, I want to enjoy life while continuing to grow."],
    similarQuestions: ["What kind of future do you want?", "Tell me about your values.", "What is important to you in life?", "What are your personal goals?", "Describe the person you want to become."],
  },
  {
    id: 31, title: "Yoga studio I often go to.", category: "Place", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Yoga studio near my house", "Quiet and peaceful atmosphere", "Clear my mind after work", "Changed exercise because of back pain", "Listen to my body", "Special recharge place"],
    keyWords: ["yoga studio", "peaceful", "back pain", "listen to my body", "recharge"],
    fullAnswer: "I usually go to a yoga studio near my house. I have been going there for quite a while, and it has become an important part of my weekly routine. The studio is quiet and peaceful. The lights are soft, and there is calm music playing in the background. Most people there are very focused on themselves, so the atmosphere feels comfortable and relaxing. One reason I like this place is that I can clear my mind there. I work as a designer, so I spend a lot of time in front of a computer. Sometimes I feel mentally tired after work, but yoga helps me feel calm again. Recently, I changed the way I exercise because I had some back pain. So now I try to listen to my body more carefully and avoid pushing myself too hard. Yoga has helped me become healthier both physically and mentally. That's why this yoga studio is special to me.",
    fullAnswerSentences: ["I usually go to a yoga studio near my house.", "I have been going there for quite a while, and it has become an important part of my weekly routine.", "The studio is quiet and peaceful.", "The lights are soft, and there is calm music playing in the background.", "Most people there are very focused on themselves, so the atmosphere feels comfortable and relaxing.", "One reason I like this place is that I can clear my mind there.", "I work as a designer, so I spend a lot of time in front of a computer.", "Sometimes I feel mentally tired after work, but yoga helps me feel calm again.", "Recently, I changed the way I exercise because I had some back pain.", "So now I try to listen to my body more carefully and avoid pushing myself too hard.", "Yoga has helped me become healthier both physically and mentally.", "That's why this yoga studio is special to me."],
    similarQuestions: ["Tell me about a place you often visit.", "Describe a place where you exercise.", "What do you usually do to stay healthy?", "Tell me about a place that helps you relax.", "What do you do after work?"],
  },
  {
    id: 32, title: "My favorite piano practice room.", category: "Place", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Favorite piano practice room", "Choose a grand piano room", "Quiet and simple space", "Piano expresses emotions", "Meaningful peaceful moment", "Important part of life"],
    keyWords: ["piano practice room", "grand piano", "quiet", "emotions", "peaceful"],
    fullAnswer: "One of my favorite places is a piano practice room that I often visit. I usually choose a room with a grand piano because I enjoy the sound and feeling of it. The room is quiet and simple, so I can focus completely on practicing. Piano is very important to me because it helps me express my emotions. When I play piano, I feel calm and focused. It almost feels like meditation. I still remember one meaningful moment. One day after work, I practiced piano for about an hour, and I suddenly felt very happy and peaceful. I realized that I really enjoy this kind of quiet time. These days, piano practice has become one of the most important parts of my life.",
    fullAnswerSentences: ["One of my favorite places is a piano practice room that I often visit.", "I usually choose a room with a grand piano because I enjoy the sound and feeling of it.", "The room is quiet and simple, so I can focus completely on practicing.", "Piano is very important to me because it helps me express my emotions.", "When I play piano, I feel calm and focused.", "It almost feels like meditation.", "I still remember one meaningful moment.", "One day after work, I practiced piano for about an hour, and I suddenly felt very happy and peaceful.", "I realized that I really enjoy this kind of quiet time.", "These days, piano practice has become one of the most important parts of my life."],
    similarQuestions: ["Describe a place that is special to you.", "Tell me about your hobby.", "What do you usually do in your free time?", "Describe a place where you spend a lot of time.", "What helps you relax?"],
  },
  {
    id: 33, title: "A quiet café I enjoy.", category: "Place", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Enjoy quiet cafés", "Need peaceful places after work", "Quiet atmosphere and soft music", "Write and think there", "Slow down and recharge"],
    keyWords: ["quiet café", "peaceful", "soft music", "notebook", "recharge"],
    fullAnswer: "I really enjoy going to quiet cafés. I usually prefer calm places because noisy places make me tired. Since I work in front of a computer all day, I sometimes need a peaceful place to relax and organize my thoughts. There is a café I like because it has a quiet atmosphere and soft music. It is not crowded, and people usually spend time reading or working quietly. When I go there, I often write in my notebook, think about my future, or simply enjoy a cup of coffee. I especially like quiet cafés because they help me slow down and recharge myself.",
    fullAnswerSentences: ["I really enjoy going to quiet cafés.", "I usually prefer calm places because noisy places make me tired.", "Since I work in front of a computer all day, I sometimes need a peaceful place to relax and organize my thoughts.", "There is a café I like because it has a quiet atmosphere and soft music.", "It is not crowded, and people usually spend time reading or working quietly.", "When I go there, I often write in my notebook, think about my future, or simply enjoy a cup of coffee.", "I especially like quiet cafés because they help me slow down and recharge myself."],
    similarQuestions: ["Describe a café you enjoy going to.", "Tell me about a place where you relax.", "Why do you like quiet places?", "What do you usually do when you need a break?", "Describe a place you visit in your free time."],
  },
  {
    id: 34, title: "A library I often visit.", category: "Place", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Library on weekends", "Part of weekend routine", "Quiet atmosphere helps me focus", "Borrow books about growth and creativity", "Discover new ideas"],
    keyWords: ["library", "weekends", "quiet", "books", "creativity", "new ideas"],
    fullAnswer: "I often go to a library on weekends. Going to the library has become part of my weekend routine. I usually stop by after exercising or spending some time outside. I enjoy the quiet atmosphere there. Everyone is focused on reading or studying, so it helps me concentrate as well. I often borrow books about writing, self-growth, or creativity because I want to become more creative in the future. Sometimes I just walk around and look at different books without any plan. I think libraries are special because they help me discover new ideas. That is one reason why I enjoy spending time there.",
    fullAnswerSentences: ["I often go to a library on weekends.", "Going to the library has become part of my weekend routine.", "I usually stop by after exercising or spending some time outside.", "I enjoy the quiet atmosphere there.", "Everyone is focused on reading or studying, so it helps me concentrate as well.", "I often borrow books about writing, self-growth, or creativity because I want to become more creative in the future.", "Sometimes I just walk around and look at different books without any plan.", "I think libraries are special because they help me discover new ideas.", "That is one reason why I enjoy spending time there."],
    similarQuestions: ["Tell me about a place you often visit.", "Describe a place where you study or read.", "What do you usually do on weekends?", "Tell me about something you do for self-improvement.", "Describe a quiet place you like."],
  },
  {
    id: 35, title: "My home is my recharge place.", category: "Place", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Home is favorite place", "Relax after work", "Piano, English, drawing, resting", "Spend time with my cat", "Alone time organizes thoughts", "Recharge place"],
    keyWords: ["home", "recharge", "cat", "alone time", "simple routine"],
    fullAnswer: "My home is probably my favorite place because it helps me recharge. I spend most of my day working, so after work, I really enjoy being at home. I like quiet and comfortable spaces because they help me relax. At home, I usually do simple things like practicing piano, studying English, drawing, or just resting. I also spend time with my cat, and that makes me feel relaxed. Sometimes, I simply stay home without any special plans. I think being alone is important because it gives me time to organize my thoughts and emotions. That is why my home feels like a recharge place to me.",
    fullAnswerSentences: ["My home is probably my favorite place because it helps me recharge.", "I spend most of my day working, so after work, I really enjoy being at home.", "I like quiet and comfortable spaces because they help me relax.", "At home, I usually do simple things like practicing piano, studying English, drawing, or just resting.", "I also spend time with my cat, and that makes me feel relaxed.", "Sometimes, I simply stay home without any special plans.", "I think being alone is important because it gives me time to organize my thoughts and emotions.", "That is why my home feels like a recharge place to me."],
    similarQuestions: ["Tell me about your home.", "Describe a place where you relax.", "What do you usually do at home?", "How do you spend time alone?", "What do you usually do on rainy days?"],
  },
  {
    id: 36, title: "My after-work routine.", category: "Routine", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Get home in the evening", "Do something healthy after work", "Exercise or yoga first", "Practice piano", "Study English and shadowing", "Routine balances body and mind"],
    keyWords: ["after work", "exercise", "yoga", "piano", "English", "balance"],
    fullAnswer: "My after-work routine is very important to me. I usually get home around the evening after work. Since I spend a lot of time sitting in front of a computer, I try to do something healthy after work. Usually, I exercise first. Sometimes I go to yoga, and sometimes I do light workouts or stretching. After that, I practice piano for about thirty minutes to an hour. These days, I also study English regularly. I listen to MP3 recordings, practice shadowing, and review my scripts. Even though I feel tired after work, I try to keep this routine because it makes me feel balanced and productive. Honestly, I think this routine helps me take care of both my body and mind.",
    fullAnswerSentences: ["My after-work routine is very important to me.", "I usually get home around the evening after work.", "Since I spend a lot of time sitting in front of a computer, I try to do something healthy after work.", "Usually, I exercise first.", "Sometimes I go to yoga, and sometimes I do light workouts or stretching.", "After that, I practice piano for about thirty minutes to an hour.", "These days, I also study English regularly.", "I listen to MP3 recordings, practice shadowing, and review my scripts.", "Even though I feel tired after work, I try to keep this routine because it makes me feel balanced and productive.", "Honestly, I think this routine helps me take care of both my body and mind."],
    similarQuestions: ["What is your after-work routine?", "Tell me about your daily routine.", "What do you usually do after work?", "How do you spend your evenings?", "What do you do to stay healthy?"],
  },
  {
    id: 37, title: "What my weekends look like.", category: "Routine", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Calm and productive weekends", "Exercise, piano, drawing, library", "Enjoy slow mornings", "Quiet café and thoughts", "Borrow books", "Recharge for new week"],
    keyWords: ["weekends", "slow mornings", "library", "quiet café", "recharge"],
    fullAnswer: "My weekends are usually calm and productive. I try to spend my weekends doing things that I enjoy. For example, I exercise, practice piano, draw, and sometimes go to the library. I especially enjoy slow mornings on weekends because I do not have to rush. I can take my time and focus on myself. Sometimes, I visit a quiet café or spend time organizing my thoughts. I also like borrowing books from the library because I enjoy learning new things. I think weekends are important because they help me recharge and prepare for a new week.",
    fullAnswerSentences: ["My weekends are usually calm and productive.", "I try to spend my weekends doing things that I enjoy.", "For example, I exercise, practice piano, draw, and sometimes go to the library.", "I especially enjoy slow mornings on weekends because I do not have to rush.", "I can take my time and focus on myself.", "Sometimes, I visit a quiet café or spend time organizing my thoughts.", "I also like borrowing books from the library because I enjoy learning new things.", "I think weekends are important because they help me recharge and prepare for a new week."],
    similarQuestions: ["What do your weekends usually look like?", "What do you usually do in your free time?", "Tell me about your weekend routine.", "How do you spend your weekends?", "Describe a relaxing day."],
  },
  {
    id: 38, title: "Why mornings matter to me.", category: "Routine", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Mornings are important", "Best energy and focus", "Quiet morning time", "Study English and organize thoughts", "Good morning improves the day", "Make mornings meaningful"],
    keyWords: ["mornings", "energy", "focus", "English", "meaningful"],
    fullAnswer: "Mornings are very important to me. I think my energy and focus are the best in the morning, so I try to use that time wisely. Usually, I like spending quiet time in the morning. Sometimes I study English, organize my thoughts, or think about ideas for drawing and writing. I realized that when I start my morning well, the rest of the day goes much better. On the other hand, if I waste my morning, I sometimes feel unmotivated all day. That is why I try to make my mornings calm and meaningful.",
    fullAnswerSentences: ["Mornings are very important to me.", "I think my energy and focus are the best in the morning, so I try to use that time wisely.", "Usually, I like spending quiet time in the morning.", "Sometimes I study English, organize my thoughts, or think about ideas for drawing and writing.", "I realized that when I start my morning well, the rest of the day goes much better.", "On the other hand, if I waste my morning, I sometimes feel unmotivated all day.", "That is why I try to make my mornings calm and meaningful."],
    similarQuestions: ["Why are mornings important to you?", "Tell me about your morning routine.", "What time of day do you like the most?", "When are you most productive?", "How do you usually start your day?"],
  },
  {
    id: 39, title: "How I study English these days.", category: "Study", set: "Set 03",
    status: "New", favorite: false, practiceCount: 0,
    flowSteps: ["Study English practically", "Focus on speaking", "Listen to MP3 and shadow", "Repeat expressions many times", "Use real-life scripts", "Practice little by little"],
    keyWords: ["English", "speaking", "MP3", "shadowing", "real-life scripts", "consistency"],
    fullAnswer: "These days, I study English in a very practical way. I mainly focus on speaking because I want to communicate naturally in English. Usually, I listen to MP3 recordings and practice shadowing. I repeat the same expressions many times because repetition helps me remember naturally. I also make my own English scripts based on my real life. Since the topics are about my daily routine and personal experiences, they feel more useful and easier to remember. Honestly, learning English is still challenging for me, but I try to practice a little every day instead of giving up. I think consistency is more important than studying too much at once.",
    fullAnswerSentences: ["These days, I study English in a very practical way.", "I mainly focus on speaking because I want to communicate naturally in English.", "Usually, I listen to MP3 recordings and practice shadowing.", "I repeat the same expressions many times because repetition helps me remember naturally.", "I also make my own English scripts based on my real life.", "Since the topics are about my daily routine and personal experiences, they feel more useful and easier to remember.", "Honestly, learning English is still challenging for me, but I try to practice a little every day instead of giving up.", "I think consistency is more important than studying too much at once."],
    similarQuestions: ["How do you study English?", "Tell me about something you are learning these days.", "What do you do for self-improvement?", "How do you practice speaking English?", "Tell me about a recent habit you made."],
  },
];

const ALL_CATEGORIES: Category[] = ["Custom", "Intro", "Hobby", "Routine", "Mindset", "Experience", "Place", "Study"];

// ─── Small Components ─────────────────────────────────────────────────────────
function ScrollToTopBottom() {
  const [visible, setVisible] = useState(false);

  useEffect(() => {
    const onScroll = () => setVisible(window.scrollY > 200);
    window.addEventListener("scroll", onScroll);
    return () => window.removeEventListener("scroll", onScroll);
  }, []);

  if (!visible) return null;

  return (
    <div className="fixed bottom-20 left-1/2 -translate-x-1/2 w-full max-w-[375px] pointer-events-none z-40">
      <div className="flex flex-col items-end gap-2 pr-5">
        <button
          onClick={() => window.scrollTo({ top: 0, behavior: "smooth" })}
          className="pointer-events-auto w-9 h-9 rounded-full bg-white border border-border shadow-sm flex items-center justify-center text-[#7A7268]"
          aria-label="Scroll to top"
        >
          <ChevronUp size={16} />
        </button>
        <button
          onClick={() => window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" })}
          className="pointer-events-auto w-9 h-9 rounded-full bg-white border border-border shadow-sm flex items-center justify-center text-[#7A7268]"
          aria-label="Scroll to bottom"
        >
          <ChevronDown size={16} />
        </button>
      </div>
    </div>
  );
}
function StatusBadge({ status }: { status: Status }) {
  const map: Record<Status, string> = {
    New: "bg-[#EEEAE3] text-[#7A7268]",
    Practicing: "bg-amber-50 text-amber-700",
    Confident: "bg-emerald-50 text-emerald-700",
  };
  return (
    <span className={`text-[11px] font-medium tracking-wide px-2 py-0.5 rounded-full ${map[status]}`}>
      {status}
    </span>
  );
}

function CategoryTag({ category }: { category: Category }) {
  const map: Record<Category, string> = {
    Custom: "text-[#B85C5C]",
    Intro: "text-[#2E7D8C]",
    Hobby: "text-[#456399]",
    Routine: "text-[#5C8C6E]",
    Mindset: "text-[#7A6AA0]",
    Experience: "text-[#A07A50]",
    Place: "text-[#9C5B7A]",
    Study: "text-[#8A7540]",
  };
  return (
    <span className={`text-[11px] font-medium uppercase tracking-widest ${map[category]}`}>
      {category}
    </span>
  );
}

// ─── Bottom Nav ───────────────────────────────────────────────────────────────

function BottomNav({ active, onTab }: { active: TabId; onTab: (t: TabId) => void }) {
  const tabs: { id: TabId; icon: typeof Home; label: string }[] = [
    { id: "home", icon: Home, label: "Home" },
    { id: "questions", icon: LayoutList, label: "Questions" },
    { id: "review", icon: BookOpen, label: "Review" },
    { id: "progress", icon: BarChart2, label: "Progress" },
  ];
  return (
    <nav className="fixed bottom-0 left-1/2 -translate-x-1/2 w-full max-w-[375px] bg-white border-t border-border z-50">
      <div className="flex">
        {tabs.map(({ id, icon: Icon, label }) => (
          <button key={id} onClick={() => onTab(id)}
            className={`flex-1 flex flex-col items-center gap-1 py-3 transition-colors ${active === id ? "text-[#456399]" : "text-[#C5C0B5]"}`}>
            <Icon size={20} strokeWidth={active === id ? 2 : 1.5} />
            <span className="text-[10px] font-medium tracking-wide">{label}</span>
          </button>
        ))}
      </div>
    </nav>
  );
}

// ─── Home Screen ──────────────────────────────────────────────────────────────

function HomeScreen({ questions, onDetail, onRandom, onTab }: {
  questions: Question[]; onDetail: (q: Question) => void; onRandom: () => void; onTab: (t: TabId) => void;
}) {
  const [activeCategory, setActiveCategory] = useState<string>("All");
  const chips = ["All", ...ALL_CATEGORIES];
  const todayQ = questions[0];
  const practiced = questions.filter(q => q.practiceCount > 0).length;
  const totalMins = Math.round(questions.reduce((s, q) => s + q.practiceCount * 0.8, 0));
  const needReview = questions.filter(q => q.status === "Practicing").length;

  return (
    <div className="pb-24 px-5 pt-8">
      <div className="mb-7">
        <p className="text-[11px] font-medium tracking-widest text-[#7A7268] uppercase mb-1">수요일 · Jun 23</p>
        <h1 className="font-display text-[26px] font-semibold text-[#17150E] leading-tight">OPIC Flow<br />Trainer</h1>
        <p className="text-sm text-[#7A7268] mt-1.5 leading-snug">Practice the order of ideas,<br />not memorized answers.</p>
      </div>

      <div className="grid grid-cols-3 gap-2.5 mb-6">
        {[
          { icon: Target, label: "Practiced", value: `${practiced}` },
          { icon: Clock, label: "Minutes", value: `${totalMins}` },
          { icon: Flame, label: "Review", value: `${needReview}` },
        ].map(({ icon: Icon, label, value }) => (
          <div key={label} className="bg-white rounded-2xl p-3.5 border border-border text-center">
            <Icon size={16} className="mx-auto mb-1.5 text-[#456399]" strokeWidth={1.5} />
            <p className="font-display text-xl font-semibold text-[#17150E]">{value}</p>
            <p className="text-[10px] text-[#7A7268] mt-0.5">{label}</p>
          </div>
        ))}
      </div>

      <div className="bg-[#456399] rounded-2xl p-5 mb-5 relative overflow-hidden">
        <div className="absolute top-0 right-0 w-32 h-32 rounded-full bg-white/5 -translate-y-8 translate-x-8" />
        <p className="text-[11px] font-medium tracking-widest text-white/60 uppercase mb-2">Today's Question</p>
        <h2 className="font-display text-[17px] font-semibold text-white leading-snug mb-3">{todayQ.title}</h2>
        <div className="flex flex-wrap gap-1.5 mb-4">
          {todayQ.flowSteps.slice(0, 3).map((s, i) => (
            <span key={i} className="text-[11px] text-white/80 bg-white/10 px-2 py-0.5 rounded-full">
              {i + 1}. {s.length > 20 ? s.slice(0, 20) + "…" : s}
            </span>
          ))}
        </div>
        <button onClick={() => onDetail(todayQ)}
          className="w-full bg-white text-[#456399] text-sm font-semibold rounded-xl py-2.5">
          Start Today's Practice
        </button>
      </div>

      <div className="flex gap-2 overflow-x-auto pb-1 mb-4 scrollbar-hide -mx-5 px-5">
        {chips.map(c => (
          <button key={c} onClick={() => setActiveCategory(c)}
            className={`flex-shrink-0 text-xs font-medium px-3.5 py-1.5 rounded-full border transition-colors ${
              activeCategory === c ? "bg-[#17150E] text-white border-[#17150E]" : "bg-white text-[#7A7268] border-border"
            }`}>{c}</button>
        ))}
      </div>

      <div className="space-y-2 mb-5">
        {questions
          .filter(q => activeCategory === "All" || q.category === activeCategory)
          .slice(0, 3)
          .map(q => (
            <button key={q.id} onClick={() => onDetail(q)}
              className="w-full bg-white rounded-2xl px-4 py-3.5 border border-border text-left flex items-center justify-between">
              <div>
                <CategoryTag category={q.category} />
                <p className="text-sm font-medium text-[#17150E] mt-0.5 leading-snug">{q.title}</p>
              </div>
              <div className="flex items-center gap-2 ml-3 flex-shrink-0">
                <StatusBadge status={q.status} />
                <ChevronRight size={14} className="text-[#C5C0B5]" />
              </div>
            </button>
          ))}
      </div>

      <div className="grid grid-cols-2 gap-3">
        <button onClick={() => onTab("questions")}
          className="bg-white border border-border rounded-2xl py-3 text-sm font-medium text-[#17150E]">
          All Questions
        </button>
        <button onClick={onRandom}
          className="bg-[#EEEAE3] rounded-2xl py-3 text-sm font-medium text-[#17150E] flex items-center justify-center gap-2">
          <Shuffle size={14} /> Random
        </button>
      </div>
    </div>
  );
}

// ─── Question List Screen ─────────────────────────────────────────────────────

function QuestionListScreen({ questions, onDetail, onToggleFav, onAdd, onExport }: {
  questions: Question[]; onDetail: (q: Question) => void;
  onToggleFav: (id: number) => void; onAdd: () => void; onExport: () => void;
}) {
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("All");
  const filters = ["All", "Set 01", "Set 02", "Set 03", ...ALL_CATEGORIES];

  const filtered = questions.filter(q => {
    const matchFilter = filter === "All" || q.set === filter || q.category === filter;
    const matchSearch = !search || q.title.toLowerCase().includes(search.toLowerCase());
    return matchFilter && matchSearch;
  });

  return (
    <div className="pb-24 pt-6">
      <div className="px-5 mb-4">
        <div className="flex items-center justify-between mb-4">
  <h2 className="font-display text-xl font-semibold text-[#17150E]">Questions</h2>
  <div className="flex items-center gap-2">
    <button onClick={onExport}
      className="flex items-center gap-1.5 bg-white border border-border text-[#7A7268] text-xs font-semibold px-3 py-1.5 rounded-full">
      <Download size={12} /> Export
    </button>
    <button onClick={onAdd}
      className="flex items-center gap-1.5 bg-[#17150E] text-white text-xs font-semibold px-3 py-1.5 rounded-full">
      <Plus size={12} /> Add Script
    </button>
  </div>
</div>
        <div className="relative">
          <Search size={15} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-[#B8B3AA]" />
          <input type="text" placeholder="Search questions…" value={search}
            onChange={e => setSearch(e.target.value)}
            className="w-full bg-white border border-border rounded-xl pl-9 pr-4 py-2.5 text-base text-[#17150E] placeholder-[#B8B3AA] outline-none focus:border-[#456399]" />
        </div>
      </div>

      <div className="flex gap-2 overflow-x-auto pb-1 mb-4 scrollbar-hide px-5">
        {filters.map(f => (
          <button key={f} onClick={() => setFilter(f)}
            className={`flex-shrink-0 text-xs font-medium px-3 py-1.5 rounded-full border transition-colors ${
              filter === f ? "bg-[#17150E] text-white border-[#17150E]" : "bg-white text-[#7A7268] border-border"
            }`}>{f}</button>
        ))}
      </div>

      <div className="space-y-2.5 px-5">
        {filtered.map(q => (
          <div key={q.id} className="bg-white rounded-2xl p-4 border border-border">
            <div className="flex items-start justify-between mb-2">
              <div className="flex items-center gap-2">
                <CategoryTag category={q.category} />
                <span className="text-[11px] text-[#C5C0B5]">{q.set}</span>
              </div>
              <div className="flex items-center gap-2">
                <StatusBadge status={q.status} />
                <button onClick={() => onToggleFav(q.id)} className="text-[#C5C0B5]">
                  <Star size={14} fill={q.favorite ? "#FBBF24" : "none"} stroke={q.favorite ? "#FBBF24" : "currentColor"} />
                </button>
              </div>
            </div>
            <button onClick={() => onDetail(q)} className="text-left w-full">
              <p className="text-[15px] font-medium text-[#17150E] mb-2.5 leading-snug">{q.title}</p>
              <div className="flex flex-wrap gap-1">
                {q.flowSteps.slice(0, 3).map((step, i) => (
                  <span key={i} className="text-[11px] text-[#7A7268] bg-[#EEEAE3] px-2 py-0.5 rounded-full">
                    {i + 1}. {step.length > 18 ? step.slice(0, 18) + "…" : step}
                  </span>
                ))}
                {q.flowSteps.length > 3 && <span className="text-[11px] text-[#B8B3AA]">+{q.flowSteps.length - 3}</span>}
              </div>
            </button>
          </div>
        ))}
        {filtered.length === 0 && (
          <div className="text-center py-12 text-[#B8B3AA] text-sm">No questions found.</div>
        )}
      </div>
    </div>
  );
}

// ─── Add Question Screen ──────────────────────────────────────────────────────

interface AddQuestionProps {
  onBack: () => void;
  onSave: (q: Question) => void;
  nextId: number;
  initialData?: Question; // 추가
}

function AddQuestionScreen({ onBack, onSave, nextId, initialData }: AddQuestionProps) {
  const [title, setTitle] = useState(initialData?.title ?? "");
  const [category, setCategory] = useState<Category>(initialData?.category ?? "Custom");
  const [set, setSet] = useState(initialData?.set ?? "Set 01");
  const [flowSteps, setFlowSteps] = useState(initialData?.flowSteps ?? ["", "", ""]);
  const [keyWords, setKeyWords] = useState(initialData?.keyWords.length ? initialData.keyWords : ["", ""]);
  const [fullAnswer, setFullAnswer] = useState(initialData?.fullAnswer ?? "");
  const [similarQuestions, setSimilarQuestions] = useState(initialData?.similarQuestions.length ? initialData.similarQuestions : ["", ""]);
  const [error, setError] = useState("");

  const updateList = (list: string[], idx: number, val: string) =>
    list.map((item, i) => i === idx ? val : item);

  const addToList = (list: string[], setList: (l: string[]) => void) =>
    setList([...list, ""]);

  const removeFromList = (list: string[], setList: (l: string[]) => void, idx: number) =>
    setList(list.filter((_, i) => i !== idx));

const handleSave = () => {
    if (!title.trim()) { setError("Please enter a question title."); return; }
    if (!fullAnswer.trim()) { setError("Please enter the full answer."); return; }
    const cleanFlow = flowSteps.filter(s => s.trim());
    if (cleanFlow.length === 0) { setError("Please add at least one flow step."); return; }

    const newQ: Question = {
      id: initialData?.id ?? nextId,           // 편집 시 기존 id 유지
      title: title.trim(),
      category,
      set,
      status: initialData?.status ?? "New",    // 편집 시 기존 status 유지
      favorite: initialData?.favorite ?? false,
      practiceCount: initialData?.practiceCount ?? 0,
      lastPracticed: initialData?.lastPracticed,
      flowSteps: cleanFlow,
      keyWords: keyWords.filter(k => k.trim()),
      fullAnswer: fullAnswer.trim(),
      fullAnswerSentences: splitSentences(fullAnswer.trim()),
      similarQuestions: similarQuestions.filter(s => s.trim()),
    };
    onSave(newQ);
  };

  const inputClass = "w-full bg-white border border-border rounded-xl px-3.5 py-2.5 text-base text-[#17150E] placeholder-[#B8B3AA] outline-none focus:border-[#456399] transition-colors";
  const labelClass = "text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-2 block";

  return (
    <div className="pb-32 pt-0">
      <div className="flex items-center justify-between px-5 pt-6 pb-4 border-b border-border">
        <button onClick={onBack} className="flex items-center gap-1.5 text-[#7A7268] text-sm">
          <ArrowLeft size={16} /> Back
        </button>
        <h2 className="font-display text-base font-semibold text-[#17150E]">
  {initialData ? "Edit Script" : "Add Script"}
</h2>
        <div className="w-12" />
      </div>

      <div className="px-5 pt-5 space-y-5">
        {/* Title */}
        <div>
          <label className={labelClass}>Question Title *</label>
          <input type="text" value={title} onChange={e => setTitle(e.target.value)}
            placeholder="e.g. Tell me about your hobbies."
            className={inputClass} />
        </div>

        {/* Category & Set row */}
        <div className="grid grid-cols-2 gap-3">
  <div>
    <label className={labelClass}>Category</label>
    <select value={category} onChange={e => setCategory(e.target.value as Category)}
      className={inputClass + " appearance-none"}>
      {ALL_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
    </select>
  </div>
  <div>
    <label className={labelClass}>Set</label>
    <select value={set} onChange={e => setSet(e.target.value)}
      className={inputClass + " appearance-none"}>
      {["Set 01", "Set 02", "Set 03", "Set 04", "Set 05"].map(s => (
        <option key={s} value={s}>{s}</option>
      ))}
    </select>
  </div>
</div>

        {/* Flow Steps */}
        <div>
          <label className={labelClass}>Flow Steps *</label>
          <div className="space-y-2">
            {flowSteps.map((step, i) => (
              <div key={i} className="flex items-center gap-2">
                <span className="flex-shrink-0 w-5 h-5 rounded-full bg-[#456399] text-white text-[10px] font-semibold flex items-center justify-center">
                  {i + 1}
                </span>
                <input type="text" value={step} onChange={e => setFlowSteps(updateList(flowSteps, i, e.target.value))}
                  placeholder={`Step ${i + 1}…`} className={inputClass + " flex-1"} />
                {flowSteps.length > 1 && (
                  <button onClick={() => removeFromList(flowSteps, setFlowSteps, i)}
                    className="text-[#C5C0B5] hover:text-red-400 flex-shrink-0">
                    <X size={14} />
                  </button>
                )}
              </div>
            ))}
            <button onClick={() => addToList(flowSteps, setFlowSteps)}
              className="flex items-center gap-1.5 text-xs text-[#456399] font-medium mt-1">
              <Plus size={12} /> Add step
            </button>
          </div>
        </div>

        {/* Key Words */}
        <div>
          <label className={labelClass}>Key Words</label>
          <div className="flex flex-wrap gap-2">
            {keyWords.map((kw, i) => (
              <div key={i} className="flex items-center gap-1 bg-[#EDF1F8] rounded-full pl-3 pr-1 py-1">
                <input type="text" value={kw} onChange={e => setKeyWords(updateList(keyWords, i, e.target.value))}
                  placeholder="keyword"
                  className="bg-transparent text-sm text-[#456399] outline-none w-20 placeholder-[#9AAAC8]" />
                <button onClick={() => removeFromList(keyWords, setKeyWords, i)}
                  className="text-[#9AAAC8] hover:text-red-400">
                  <X size={11} />
                </button>
              </div>
            ))}
            <button onClick={() => addToList(keyWords, setKeyWords)}
              className="flex items-center gap-1 text-xs text-[#456399] font-medium bg-[#EDF1F8] rounded-full px-3 py-1">
              <Plus size={11} /> Add
            </button>
          </div>
        </div>

        {/* Full Answer */}
        <div>
          <label className={labelClass}>Full Answer *</label>
          <textarea value={fullAnswer} onChange={e => setFullAnswer(e.target.value)}
            placeholder="Write your full answer here. It will be split into sentences automatically for shadowing practice."
            rows={6}
            className={inputClass + " resize-none leading-relaxed"} />
          <p className="text-[11px] text-[#B8B3AA] mt-1.5">
            Sentences are split automatically at each period.
          </p>
        </div>

        {/* Similar Questions */}
        <div>
          <label className={labelClass}>Similar Questions</label>
          <div className="space-y-2">
            {similarQuestions.map((sq, i) => (
              <div key={i} className="flex items-center gap-2">
                <span className="text-[#C5C0B5] flex-shrink-0 text-sm">·</span>
                <input type="text" value={sq} onChange={e => setSimilarQuestions(updateList(similarQuestions, i, e.target.value))}
                  placeholder={`Similar question ${i + 1}…`} className={inputClass + " flex-1"} />
                {similarQuestions.length > 1 && (
                  <button onClick={() => removeFromList(similarQuestions, setSimilarQuestions, i)}
                    className="text-[#C5C0B5] hover:text-red-400 flex-shrink-0">
                    <X size={14} />
                  </button>
                )}
              </div>
            ))}
            <button onClick={() => addToList(similarQuestions, setSimilarQuestions)}
              className="flex items-center gap-1.5 text-xs text-[#456399] font-medium mt-1">
              <Plus size={12} /> Add similar question
            </button>
          </div>
        </div>
      </div>

      {/* Fixed bottom CTA */}
      <div className="fixed bottom-0 left-1/2 -translate-x-1/2 w-full max-w-[375px] bg-[#F7F5F1] border-t border-border px-5 py-4">
        {error && <p className="text-xs text-red-500 mb-2 text-center">{error}</p>}
        <button onClick={handleSave}
          className="w-full bg-[#17150E] text-white text-sm font-semibold rounded-2xl py-3.5">
          Save Script
        </button>
      </div>
    </div>
  );
}

// ─── Question Detail Screen ───────────────────────────────────────────────────

function QuestionDetailScreen({ question: q, onBack, onPractice, onShadowing, onToggleFav, onStatusChange, onEdit }: {
  question: Question; onBack: () => void; onPractice: () => void; onShadowing: () => void;
  onToggleFav: (id: number) => void; onStatusChange: (id: number, s: Status) => void;
  onEdit: () => void; // 추가
}) {
  const [showAnswer, setShowAnswer] = useState(false);
  const [showSimilar, setShowSimilar] = useState(false);
  const [flowOnly, setFlowOnly] = useState(false);
  const statuses: Status[] = ["New", "Practicing", "Confident"];

  return (
    <div className="pb-28 pt-0">
      <div className="flex items-center justify-between px-5 pt-6 pb-4">
  <button onClick={onBack} className="flex items-center gap-1.5 text-[#7A7268] text-sm">
    <ArrowLeft size={16} /> Back
  </button>
  <div className="flex items-center gap-3">
    <button onClick={onEdit}
      className="text-xs text-[#7A7268] border border-border rounded-full px-3 py-1">
      Edit
    </button>
    <button onClick={() => onToggleFav(q.id)}>
      <Star size={18} fill={q.favorite ? "#FBBF24" : "none"} stroke={q.favorite ? "#FBBF24" : "#C5C0B5"} />
    </button>
  </div>
</div>

      <div className="px-5 space-y-4 pb-10">
        <div>
          <CategoryTag category={q.category} />
          <h1 className="font-display text-[22px] font-semibold text-[#17150E] mt-1.5 leading-snug">{q.title}</h1>
          <p className="text-xs text-[#7A7268] mt-2 italic">First, remember the order of ideas.</p>
        </div>

        <div className="flex gap-2">
          {statuses.map(s => (
            <button key={s} onClick={() => onStatusChange(q.id, s)}
              className={`flex-1 text-xs py-1.5 rounded-lg border transition-colors font-medium ${
                q.status === s
                  ? s === "New" ? "bg-[#EEEAE3] text-[#17150E] border-transparent"
                  : s === "Practicing" ? "bg-amber-50 text-amber-700 border-amber-200"
                  : "bg-emerald-50 text-emerald-700 border-emerald-200"
                  : "bg-white text-[#B8B3AA] border-border"
              }`}>{s}</button>
          ))}
        </div>

        <div className="bg-white rounded-2xl border border-border overflow-hidden">
          <button onClick={() => setShowSimilar(v => !v)}
            className="w-full flex items-center justify-between px-4 py-3.5 text-left">
            <span className="text-sm font-medium text-[#17150E]">Similar Questions</span>
            {showSimilar ? <ChevronUp size={15} className="text-[#B8B3AA]" /> : <ChevronDown size={15} className="text-[#B8B3AA]" />}
          </button>
          {showSimilar && (
            <div className="border-t border-border px-4 py-3 space-y-2">
              {q.similarQuestions.map((sq, i) => (
                <p key={i} className="text-sm text-[#7A7268] leading-snug flex gap-2">
                  <span className="text-[#C5C0B5]">·</span> {sq}
                </p>
              ))}
            </div>
          )}
        </div>

        <div className="bg-white rounded-2xl border border-border p-4">
          <h3 className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-3">Flow Steps</h3>
          <div className="space-y-2.5">
            {q.flowSteps.map((step, i) => (
              <div key={i} className="flex items-start gap-3">
                <span className="flex-shrink-0 w-5 h-5 rounded-full bg-[#456399] text-white text-[10px] font-semibold flex items-center justify-center mt-0.5">
                  {i + 1}
                </span>
                <p className="text-sm text-[#17150E] leading-snug pt-0.5">{step}</p>
              </div>
            ))}
          </div>
        </div>

        {!flowOnly && q.keyWords.length > 0 && (
          <div className="bg-white rounded-2xl border border-border p-4">
            <h3 className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-3">Key Words</h3>
            <div className="flex flex-wrap gap-2">
              {q.keyWords.map(kw => (
                <span key={kw} className="text-sm text-[#456399] bg-[#EDF1F8] px-3 py-1 rounded-full font-medium">{kw}</span>
              ))}
            </div>
          </div>
        )}

        {!flowOnly && (
          <div className="bg-white rounded-2xl border border-border overflow-hidden">
            <button onClick={() => setShowAnswer(v => !v)}
              className="w-full flex items-center justify-between px-4 py-3.5 text-left">
              <span className="text-sm font-medium text-[#17150E]">Full Answer</span>
              {showAnswer ? <EyeOff size={15} className="text-[#B8B3AA]" /> : <Eye size={15} className="text-[#456399]" />}
            </button>
            {showAnswer && (
              <div className="border-t border-border px-4 py-3">
                <p className="text-sm text-[#17150E] leading-relaxed">{q.fullAnswer}</p>
              </div>
            )}
          </div>
        )}

        <button onClick={() => setFlowOnly(v => !v)}
          className="w-full text-xs text-[#7A7268] py-1 underline underline-offset-2 text-center">
          {flowOnly ? "Show full details" : "Flow Only mode"}
        </button>
      </div>

      <div className="fixed bottom-0 left-1/2 -translate-x-1/2 w-full max-w-[375px] bg-[#F7F5F1] border-t border-border px-5 py-4 space-y-2.5">
        <button onClick={onPractice}
          className="w-full bg-[#17150E] text-white text-sm font-semibold rounded-2xl py-3.5">
          Start 60s Practice
        </button>
        <div className="grid grid-cols-2 gap-2.5">
          <button onClick={() => setShowAnswer(true)}
            className="bg-white border border-border rounded-2xl py-2.5 text-sm font-medium text-[#17150E] flex items-center justify-center gap-1.5">
            <Eye size={14} /> Full Answer
          </button>
          <button onClick={onShadowing}
            className="bg-[#EDF1F8] rounded-2xl py-2.5 text-sm font-medium text-[#456399] flex items-center justify-center gap-1.5">
            <RotateCcw size={14} /> Shadowing
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Speaking Practice Screen ─────────────────────────────────────────────────

function SpeakingPracticeScreen({ question: q, onBack, onSave }: {
  question: Question; onBack: () => void; onSave: () => void;
}) {
  const [elapsed, setElapsed] = useState(0);
  const [isRecording, setIsRecording] = useState(false);
  const [isDone, setIsDone] = useState(false);
  const [checks, setChecks] = useState({ flow: false, over40: false, again: false });
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    if (isRecording) {
      timerRef.current = setInterval(() => setElapsed(e => e + 1), 1000);
    }
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [isRecording]);

  const startRecording = () => { setElapsed(0); setIsRecording(true); setIsDone(false); };
  const stopRecording = () => { setIsRecording(false); setIsDone(true); };

  const timerColor = elapsed < 40 ? "#7A7268" : elapsed < 60 ? "#456399" : "#C07A30";
  const timerMsg = !isRecording && !isDone ? "Tap the mic to begin"
    : elapsed < 40 ? "Keep going…" : elapsed < 60 ? "Good range ✓" : "Wrap up now";

  return (
    <div className="min-h-screen flex flex-col px-5 pt-6">
      <button onClick={onBack} className="flex items-center gap-1.5 text-[#7A7268] text-sm mb-6">
        <ArrowLeft size={16} /> Back
      </button>
      <div className="mb-6">
        <p className="text-[11px] font-medium tracking-widest text-[#7A7268] uppercase mb-1">Speaking</p>
        <h2 className="font-display text-lg font-semibold text-[#17150E] leading-snug">{q.title}</h2>
      </div>
      <div className="grid grid-cols-1 gap-1.5 mb-8">
        {q.flowSteps.map((step, i) => (
          <div key={i} className="flex items-center gap-2.5 bg-white rounded-xl px-3.5 py-2.5 border border-border">
            <span className="flex-shrink-0 text-[10px] font-semibold text-[#456399] w-4 font-mono-data">{i + 1}.</span>
            <p className="text-[13px] text-[#17150E]">{step}</p>
          </div>
        ))}
      </div>
      <div className="flex flex-col items-center mb-8">
        <span className="font-mono-data text-6xl font-medium tracking-tight transition-colors" style={{ color: timerColor }}>
          {FmtTime(elapsed)}
        </span>
        <p className="text-sm mt-2 transition-all" style={{ color: timerColor }}>{timerMsg}</p>
      </div>
      {!isDone && (
        <div className="flex justify-center gap-5 mb-8">
          {!isRecording ? (
            <button onClick={startRecording}
              className="w-20 h-20 rounded-full bg-[#17150E] text-white flex items-center justify-center shadow-lg active:scale-95 transition-transform">
              <Mic size={28} strokeWidth={1.5} />
            </button>
          ) : (
            <button onClick={stopRecording}
              className="w-20 h-20 rounded-full bg-red-500 text-white flex items-center justify-center shadow-lg active:scale-95 transition-transform">
              <Square size={24} fill="white" />
            </button>
          )}
        </div>
      )}
      {isDone && (
        <div className="bg-white rounded-2xl border border-border p-5">
          <p className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-4">Self Check</p>
          {([
            { key: "flow" as const, label: "I followed the flow" },
            { key: "over40" as const, label: "I spoke over 40 seconds" },
            { key: "again" as const, label: "I need to practice again" },
          ]).map(({ key, label }) => (
            <button key={key} onClick={() => setChecks(c => ({ ...c, [key]: !c[key] }))}
              className="w-full flex items-center gap-3 py-2.5 border-b border-border last:border-0 text-left">
              <div className={`w-5 h-5 rounded-md border flex items-center justify-center transition-colors ${checks[key] ? "bg-[#17150E] border-[#17150E]" : "border-border"}`}>
                {checks[key] && <Check size={12} className="text-white" strokeWidth={2.5} />}
              </div>
              <span className="text-sm text-[#17150E]">{label}</span>
            </button>
          ))}
          <div className="flex gap-2.5 mt-4">
            <button onClick={startRecording}
              className="flex-1 bg-[#EEEAE3] rounded-xl py-2.5 text-sm font-medium text-[#17150E]">
              Try Again
            </button>
            <button onClick={onSave}
              className="flex-1 bg-[#17150E] text-white rounded-xl py-2.5 text-sm font-semibold">
              Save Practice
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Shadowing Screen ─────────────────────────────────────────────────────────

function ShadowingScreen({ question: q, onBack, onSpeakAgain }: {
  question: Question; onBack: () => void; onSpeakAgain: () => void;
}) {
  const [idx, setIdx] = useState(0);
  const [showAnswer, setShowAnswer] = useState(true);
  const sentences = q.fullAnswerSentences;

  return (
    <div className="min-h-screen flex flex-col px-5 pt-6 pb-10">
      <button onClick={onBack} className="flex items-center gap-1.5 text-[#7A7268] text-sm mb-6">
        <ArrowLeft size={16} /> Back
      </button>
      <div className="mb-4">
        <p className="text-[11px] font-medium tracking-widest text-[#7A7268] uppercase mb-1">Shadowing</p>
        <h2 className="font-display text-lg font-semibold text-[#17150E] leading-snug">{q.title}</h2>
      </div>
      <div className="flex gap-1.5 mb-6">
        {sentences.map((_, i) => (
          <div key={i} className={`h-1 flex-1 rounded-full transition-colors ${i === idx ? "bg-[#456399]" : i < idx ? "bg-[#456399]/30" : "bg-[#EEEAE3]"}`} />
        ))}
      </div>
      <div className="flex-1 flex flex-col">
        <div className="bg-white rounded-2xl border border-border p-6 mb-4 flex-1 flex flex-col justify-center">
          <p className="text-[11px] font-medium tracking-widest text-[#7A7268] uppercase mb-3">{idx + 1} / {sentences.length}</p>
          {showAnswer
            ? <p className="font-display text-[19px] font-medium text-[#17150E] leading-relaxed">{sentences[idx]}</p>
            : <div className="h-8 bg-[#EEEAE3] rounded-lg" />}
        </div>
        {q.keyWords.length > 0 && (
          <div className="mb-5">
            <p className="text-[11px] font-medium tracking-widest text-[#7A7268] uppercase mb-2">Key Words</p>
            <div className="flex flex-wrap gap-2">
              {q.keyWords.map(kw => (
                <span key={kw} className="text-sm text-[#456399] bg-[#EDF1F8] px-3 py-1 rounded-full font-medium">{kw}</span>
              ))}
            </div>
          </div>
        )}
        <div className="flex items-center justify-between mb-4">
          <button onClick={() => setIdx(i => Math.max(0, i - 1))} disabled={idx === 0}
            className="w-12 h-12 rounded-full bg-white border border-border flex items-center justify-center disabled:opacity-30">
            <SkipBack size={18} className="text-[#17150E]" />
          </button>
          <button onClick={() => setShowAnswer(v => !v)}
            className="flex items-center gap-2 px-4 h-10 rounded-full bg-[#EEEAE3] text-sm font-medium text-[#17150E]">
            {showAnswer ? <EyeOff size={14} /> : <Eye size={14} />}
            {showAnswer ? "Hide" : "Show"}
          </button>
          <button onClick={() => setIdx(i => Math.min(sentences.length - 1, i + 1))} disabled={idx === sentences.length - 1}
            className="w-12 h-12 rounded-full bg-white border border-border flex items-center justify-center disabled:opacity-30">
            <SkipForward size={18} className="text-[#17150E]" />
          </button>
        </div>
        <button onClick={onSpeakAgain}
          className="w-full bg-[#17150E] text-white text-sm font-semibold rounded-2xl py-3.5">
          Speak Again Without Answer
        </button>
      </div>
    </div>
  );
}

// ─── Review Screen ────────────────────────────────────────────────────────────

function ReviewScreen({ questions, onDetail }: { questions: Question[]; onDetail: (q: Question) => void }) {
  const [activeTab, setActiveTab] = useState<"difficult" | "favorites" | "recent">("difficult");
  const difficult = questions.filter(q => q.status === "Practicing");
  const favorites = questions.filter(q => q.favorite);
  const recent = [...questions].filter(q => q.lastPracticed).slice(0, 5);
  const displayList = activeTab === "difficult" ? difficult : activeTab === "favorites" ? favorites : recent;

  return (
    <div className="pb-24 pt-6 px-5">
      <h2 className="font-display text-xl font-semibold text-[#17150E] mb-5">Review Notes</h2>
      <div className="flex gap-1 bg-[#EEEAE3] p-1 rounded-xl mb-5">
        {([
          { key: "difficult" as const, label: `Difficult (${difficult.length})` },
          { key: "favorites" as const, label: `Favorites (${favorites.length})` },
          { key: "recent" as const, label: `Recent (${recent.length})` },
        ]).map(({ key, label }) => (
          <button key={key} onClick={() => setActiveTab(key)}
            className={`flex-1 text-xs font-medium py-1.5 rounded-lg transition-all ${activeTab === key ? "bg-white text-[#17150E] shadow-sm" : "text-[#7A7268]"}`}>
            {label}
          </button>
        ))}
      </div>
      <div className="space-y-2.5">
        {displayList.map(q => (
          <button key={q.id} onClick={() => onDetail(q)}
            className="w-full bg-white rounded-2xl px-4 py-4 border border-border text-left">
            <div className="flex items-start justify-between mb-1.5">
              <CategoryTag category={q.category} />
              <div className="flex items-center gap-2">
                {q.lastPracticed && <span className="text-[11px] text-[#B8B3AA]">{q.lastPracticed}</span>}
                <StatusBadge status={q.status} />
              </div>
            </div>
            <p className="text-sm font-medium text-[#17150E] leading-snug mb-2">{q.title}</p>
            <div className="flex flex-wrap gap-1">
              {q.flowSteps.slice(0, 2).map((s, i) => (
                <span key={i} className="text-[11px] text-[#7A7268] bg-[#EEEAE3] px-2 py-0.5 rounded-full">
                  {s.length > 20 ? s.slice(0, 20) + "…" : s}
                </span>
              ))}
            </div>
          </button>
        ))}
        {displayList.length === 0 && (
          <div className="text-center py-12 text-[#B8B3AA] text-sm">Nothing here yet.</div>
        )}
      </div>
      {displayList.length > 0 && (
        <button className="w-full mt-5 bg-[#EEEAE3] rounded-2xl py-3 text-sm font-medium text-[#17150E]">
          Review All ({displayList.length})
        </button>
      )}
    </div>
  );
}

// ─── Progress Screen ──────────────────────────────────────────────────────────

function ProgressScreen({ questions, practiceLog }: { questions: Question[]; practiceLog: PracticeLog }) {
  // 최근 7일 (오늘 포함, 과거 순으로 정렬)
  const last7 = Array.from({ length: 7 }, (_, i) => 6 - i); // [6,5,4,3,2,1,0] daysAgo
  const weekDays = last7.map(d => shortDayLabel(d));
  const weekMins = last7.map(d => Math.round(practiceLog[dateKeyFrom(d)] ?? 0));
  const streak = last7.map(d => (practiceLog[dateKeyFrom(d)] ?? 0) > 0);
  const maxMins = Math.max(...weekMins, 1);

  // 연속 일수 계산: 오늘부터 거꾸로 기록 있는 날 카운트
  let streakCount = 0;
  for (let i = 0; ; i++) {
    const key = dateKeyFrom(i);
    if ((practiceLog[key] ?? 0) > 0) streakCount++;
    else break;
  }

  const totalMins = Math.round(Object.values(practiceLog).reduce((s, v) => s + v, 0));
  const totalDone = questions.filter(q => q.status === "Confident").length;

  const catColors: Record<Category, string> = {
    Custom: "#B85C5C", Intro: "#2E7D8C", Hobby: "#456399", Routine: "#5C8C6E",
    Mindset: "#7A6AA0", Experience: "#A07A50", Place: "#9C5B7A", Study: "#8A7540",
  };
  const categoryStats = ALL_CATEGORIES.map(cat => ({
    name: cat,
    done: questions.filter(q => q.category === cat && q.status === "Confident").length,
    total: questions.filter(q => q.category === cat).length,
    color: catColors[cat],
  })).filter(c => c.total > 0);

  return (
    <div className="pb-24 pt-6 px-5">
      <h2 className="font-display text-xl font-semibold text-[#17150E] mb-5">Progress</h2>
      <div className="grid grid-cols-3 gap-2.5 mb-5">
        {[
          { icon: Flame, label: "Streak", value: `${streakCount}d` },
          { icon: Clock, label: "Minutes", value: `${totalMins}` },
          { icon: Target, label: "Confident", value: `${totalDone}` },
        ].map(({ icon: Icon, label, value }) => (
          <div key={label} className="bg-white rounded-2xl p-3.5 border border-border text-center">
            <Icon size={16} className="mx-auto mb-1.5 text-[#456399]" strokeWidth={1.5} />
            <p className="font-display text-xl font-semibold text-[#17150E]">{value}</p>
            <p className="text-[10px] text-[#7A7268] mt-0.5">{label}</p>
          </div>
        ))}
      </div>

      <div className="bg-white rounded-2xl border border-border p-4 mb-4">
        <p className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-4">This Week</p>
        <div className="flex items-end justify-between gap-2 h-20 mb-2">
          {weekMins.map((mins, i) => (
            <div key={i} className="flex-1 flex flex-col items-center justify-end">
              <div className="w-full rounded-t-md transition-all"
                style={{ height: mins > 0 ? `${(mins / maxMins) * 100}%` : "4px", backgroundColor: mins > 0 ? "#456399" : "#EEEAE3", minHeight: "4px" }} />
            </div>
          ))}
        </div>
        <div className="flex justify-between">
          {weekDays.map((d, i) => <span key={i} className="flex-1 text-center text-[10px] text-[#B8B3AA] font-medium">{d}</span>)}
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-border p-4 mb-4">
        <p className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-3">Daily Streak</p>
        <div className="flex justify-between">
          {streak.map((active, i) => (
            <div key={i} className="flex flex-col items-center gap-1.5">
              <div className={`w-8 h-8 rounded-full flex items-center justify-center ${active ? "bg-[#456399]" : "bg-[#EEEAE3]"}`}>
                {active && <Check size={12} className="text-white" strokeWidth={2.5} />}
              </div>
              <span className="text-[10px] text-[#B8B3AA]">{weekDays[i]}</span>
            </div>
          ))}
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-border p-4">
        <p className="text-xs font-semibold uppercase tracking-widest text-[#7A7268] mb-4">By Category</p>
        <div className="space-y-4">
          {categoryStats.map(({ name, done, total, color }) => (
            <div key={name}>
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-sm font-medium text-[#17150E]">{name}</span>
                <span className="text-xs text-[#7A7268]">{done}/{total}</span>
              </div>
              <div className="h-1.5 bg-[#EEEAE3] rounded-full overflow-hidden">
                <div className="h-full rounded-full transition-all"
                  style={{ width: `${total > 0 ? (done / total) * 100 : 0}%`, backgroundColor: color }} />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── App Root ─────────────────────────────────────────────────────────────────

export default function App() {
  const [screen, setScreen] = useState<Screen>("home");
  const [activeTab, setActiveTab] = useState<TabId>("home");
  const [questions, setQuestions] = useState<Question[]>(() => {
    const saved = localStorage.getItem("opic-questions");
    return saved ? JSON.parse(saved) : INITIAL_QUESTIONS;
  });

  useEffect(() => {
    localStorage.setItem("opic-questions", JSON.stringify(questions));
  }, [questions]);

  const [practiceLog, setPracticeLog] = useState<PracticeLog>(() => {
    const saved = localStorage.getItem("opic-practice-log");
    return saved ? JSON.parse(saved) : {};
  });

  useEffect(() => {
    localStorage.setItem("opic-practice-log", JSON.stringify(practiceLog));
  }, [practiceLog]);

  const [selectedQ, setSelectedQ] = useState<Question | null>(null);

  const navigateTo = (s: Screen, q?: Question) => {
    if (q) setSelectedQ(q);
    setScreen(s);
    if (["home", "questions", "review", "progress"].includes(s)) setActiveTab(s as TabId);
  };

  const handleTab = (t: TabId) => { setActiveTab(t); setScreen(t); };

  const handleRandom = () => {
    const q = questions[Math.floor(Math.random() * questions.length)];
    navigateTo("detail", q);
  };

  const handleToggleFav = (id: number) => {
    setQuestions(qs => qs.map(q => q.id === id ? { ...q, favorite: !q.favorite } : q));
    setSelectedQ(q => q?.id === id ? { ...q, favorite: !q.favorite } : q);
  };

  const handleStatusChange = (id: number, status: Status) => {
    setQuestions(qs => qs.map(q => q.id === id ? { ...q, status } : q));
    setSelectedQ(q => q?.id === id ? { ...q, status } : q);
  };

  const handleSavePractice = () => {
  if (selectedQ) {
    setQuestions(qs => qs.map(q => q.id === selectedQ.id
      ? { ...q, practiceCount: q.practiceCount + 1, lastPracticed: shortDateLabel() } : q));

    const key = todayKey();
    setPracticeLog(log => ({ ...log, [key]: (log[key] ?? 0) + 0.8 }));
  }
  navigateTo("detail");
};

const handleEditQuestion = (updatedQ: Question) => {
  setQuestions(qs => qs.map(q => q.id === updatedQ.id ? updatedQ : q));
  setSelectedQ(updatedQ);
  navigateTo("detail");
};
  const handleAddQuestion = (newQ: Question) => {
    setQuestions(qs => [...qs, newQ]);
    navigateTo("questions");
  };
const handleExportExcel = () => {
  // 질문 데이터를 엑셀 행 형태로 변환
  const rows = questions.map(q => ({
    ID: q.id,
    Title: q.title,
    Category: q.category,
    Set: q.set,
    Status: q.status,
    Favorite: q.favorite ? "Yes" : "No",
    "Practice Count": q.practiceCount,
    "Last Practiced": q.lastPracticed ?? "",
    "Flow Steps": q.flowSteps.join(" → "),
    "Key Words": q.keyWords.join(", "),
    "Full Answer": q.fullAnswer,
    "Similar Questions": q.similarQuestions.join(" / "),
  }));

  const worksheet = XLSX.utils.json_to_sheet(rows);

  // 열 너비 보기 좋게 조정
  worksheet["!cols"] = [
    { wch: 5 }, { wch: 35 }, { wch: 12 }, { wch: 8 }, { wch: 12 },
    { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 50 }, { wch: 30 },
    { wch: 60 }, { wch: 50 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Questions");

  XLSX.writeFile(workbook, "opic_questions.xlsx");
};
  const isMainTab = ["home", "questions", "review", "progress"].includes(screen);
  const currentQ = selectedQ ?? questions[0];

  return (
    <div className="min-h-screen bg-background flex justify-center">
      <div className="w-full max-w-[375px] relative bg-background min-h-screen">
        {screen === "home" && (
          <HomeScreen questions={questions} onDetail={q => navigateTo("detail", q)} onRandom={handleRandom} onTab={handleTab} />
        )}
        {screen === "questions" && (
          <QuestionListScreen questions={questions} onDetail={q => navigateTo("detail", q)}
            onToggleFav={handleToggleFav} onAdd={() => navigateTo("add")}
            onExport={handleExportExcel} />
        )}
        {screen === "add" && (
          <AddQuestionScreen onBack={() => navigateTo("questions")}
            onSave={handleAddQuestion} nextId={questions.length + 1} />
        )}
        {screen === "detail" && (
          <QuestionDetailScreen question={currentQ} onBack={() => navigateTo(activeTab)}
            onPractice={() => navigateTo("practice")} onShadowing={() => navigateTo("shadowing")}
            onToggleFav={handleToggleFav} onStatusChange={handleStatusChange}
            onEdit={() => navigateTo("edit")} />  // 추가
        )}

        {screen === "edit" && (
          <AddQuestionScreen onBack={() => navigateTo("detail")}
            onSave={handleEditQuestion}
            nextId={questions.length + 1}
            initialData={currentQ} />  // 추가
        )}
        {screen === "practice" && (
          <SpeakingPracticeScreen question={currentQ} onBack={() => navigateTo("detail")} onSave={handleSavePractice} />
        )}
        {screen === "shadowing" && (
          <ShadowingScreen question={currentQ} onBack={() => navigateTo("detail")} onSpeakAgain={() => navigateTo("practice")} />
        )}
        {screen === "review" && (
          <ReviewScreen questions={questions} onDetail={q => navigateTo("detail", q)} />
        )}
        {screen === "progress" && <ProgressScreen questions={questions} practiceLog={practiceLog} />}

        {isMainTab && <BottomNav active={activeTab} onTab={handleTab} />}
         <ScrollToTopBottom />
      </div>
    </div>
  );
}
